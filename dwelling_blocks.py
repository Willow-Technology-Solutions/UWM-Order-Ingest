"""Dwelling Blocks API client and ERP row mapping.

Auth: POST /auth/token with client_id + client_secret from .env
List: GET /vendors/orders
Search: POST /vendors/orders/search
Detail: GET /vendors/orders/{orderAssignmentId} for contacts and properties

ContactType includes Borrower and AccessContact. There is no CoBorrower type.
Extra borrowers are additional contacts with contactType == "Borrower".

Docs:
  https://dwelling-blocks.readme.io/reference/getvendororders
  https://dwelling-blocks.readme.io/reference/searchcompanyassignments
  https://dwelling-blocks.readme.io/reference/getvendororder
"""

from __future__ import annotations

import os
import re
from typing import Any, Iterator

import requests

from helpers import DNS_SERVER, _createDnsPinnedSession

API_BASE = "https://api.dwellingblocks.com"
DEFAULT_PAGE_SIZE = 100

# Filters used by https://app.dwellingblocks.com/assignments/in-progress
# Accepted + (no report uploaded OR open revision due)
IN_PROGRESS_FILTERS: list[dict[str, str]] = [
    {"assignmentStatus": "Accepted", "reportUploadedDate": "!"},
    {"assignmentStatus": "Accepted", "openRevisionDueDate": "*"},
]

VIEW_FILTERS: dict[str, list[dict[str, str]] | None] = {
    "in-progress": IN_PROGRESS_FILTERS,
    "all": None,
}


class DwellingBlocksClient:
    """Thin client for Dwelling Blocks Vendor assignment endpoints."""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        *,
        base_url: str = API_BASE,
        timeout: float = 30,
    ) -> None:
        self.client_id = client_id
        self.client_secret = client_secret
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self._access_token: str | None = None
        self._session = _createDnsPinnedSession(dns_server=DNS_SERVER)

    def getAccessToken(self, *, force: bool = False) -> str:
        if self._access_token and not force:
            return self._access_token

        response = self._session.post(
            f"{self.base_url}/auth/token",
            headers={"Content-Type": "application/json"},
            json={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            },
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        token = data.get("access_token")
        if not token:
            raise RuntimeError(f"No access_token in auth response: {data}")
        self._access_token = token
        return token

    def _authHeaders(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self.getAccessToken()}"}

    def getVendorOrdersPage(
        self, *, offset: int = 0, num_results: int = DEFAULT_PAGE_SIZE
    ) -> dict[str, Any]:
        """GET /vendors/orders — list of order assignments."""
        response = self._session.get(
            f"{self.base_url}/vendors/orders",
            headers=self._authHeaders(),
            params={"offset": offset, "numResults": num_results},
            timeout=self.timeout,
        )
        response.raise_for_status()
        return response.json()

    def searchVendorOrdersPage(
        self,
        *,
        offset: int = 0,
        num_results: int = DEFAULT_PAGE_SIZE,
        query: str | None = None,
        sort: str | None = "createdOn",
        direction: str | None = "Desc",
        filters: list[dict[str, str]] | None = None,
        item: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """POST /vendors/orders/search — company assignment search (UI/CSV source)."""
        body: dict[str, Any] = {
            "offset": offset,
            "numResults": num_results,
            "item": item if item is not None else {},
            "filters": filters if filters is not None else [],
        }
        if query is not None:
            body["query"] = query
        if sort is not None:
            body["sort"] = sort
        if direction is not None:
            body["direction"] = direction

        response = self._session.post(
            f"{self.base_url}/vendors/orders/search",
            headers={**self._authHeaders(), "Content-Type": "application/json"},
            json=body,
            timeout=self.timeout,
        )
        response.raise_for_status()
        return response.json()

    def getVendorOrder(self, order_assignment_id: int) -> dict[str, Any]:
        """GET /vendors/orders/{orderAssignmentId} — assigned order details."""
        response = self._session.get(
            f"{self.base_url}/vendors/orders/{order_assignment_id}",
            headers=self._authHeaders(),
            timeout=self.timeout,
        )
        response.raise_for_status()
        return response.json()

    def getVendorOrderTransfers(
        self, order_assignment_id: int
    ) -> list[dict[str, Any]]:
        """GET /vendors/orders/{orderAssignmentId}/transfers — payout transfers."""
        response = self._session.get(
            f"{self.base_url}/vendors/orders/{order_assignment_id}/transfers",
            headers=self._authHeaders(),
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        return data if isinstance(data, list) else []

    def iterVendorOrders(
        self,
        *,
        page_size: int = DEFAULT_PAGE_SIZE,
        max_results: int | None = None,
        use_search: bool = True,
        query: str | None = None,
        sort: str | None = "createdOn",
        direction: str | None = "Desc",
        filters: list[dict[str, str]] | None = None,
    ) -> Iterator[dict[str, Any]]:
        """Yield assignment items, paging until exhausted or max_results."""
        offset = 0
        yielded = 0

        while True:
            remaining = None if max_results is None else max_results - yielded
            if remaining is not None and remaining <= 0:
                break

            num_results = page_size if remaining is None else min(page_size, remaining)
            if use_search:
                page = self.searchVendorOrdersPage(
                    offset=offset,
                    num_results=num_results,
                    query=query,
                    sort=sort,
                    direction=direction,
                    filters=filters,
                )
            else:
                page = self.getVendorOrdersPage(
                    offset=offset, num_results=num_results
                )

            items = page.get("items") or []
            if not items:
                break

            for item in items:
                yield item
                yielded += 1
                if max_results is not None and yielded >= max_results:
                    return

            total = page.get("totalCount")
            offset += len(items)
            if total is not None and offset >= total:
                break
            if len(items) < num_results:
                break


def contactsOfType(
    contacts: list[dict[str, Any]] | None,
    contact_type: str,
    *,
    sort_by_id_desc: bool = False,
) -> list[dict[str, Any]]:
    """Return contacts matching ContactType (e.g. Borrower, AccessContact)."""
    matched = [
        contact
        for contact in contacts or []
        if (contact.get("contactType") or "").strip() == contact_type
    ]
    if sort_by_id_desc:
        matched.sort(key=lambda c: int(c.get("id") or 0), reverse=True)
    return matched


def propertyValueByLabel(
    properties: list[dict[str, Any]] | None, label: str
) -> str | None:
    for prop in properties or []:
        if (prop.get("label") or "").strip() != label:
            continue
        value = prop.get("value") or {}
        selected = value.get("selectedValue") or value.get("value")
        if selected is not None and str(selected).strip() != "":
            return str(selected)
    return None


def vendorPayoutCents(transfers: list[dict[str, Any]] | None) -> int | None:
    """Sum Vendor-type transfer totals (API amounts are integer cents)."""
    cents = 0
    found = False
    for transfer in transfers or []:
        if (transfer.get("transferType") or "").strip() != "Vendor":
            continue
        found = True
        if transfer.get("total") is not None:
            cents += int(transfer["total"])
        else:
            amount = int(transfer.get("amount") or 0)
            reversed_amount = int(transfer.get("reversedAmount") or 0)
            cents += amount - reversed_amount
    return cents if found else None


def formatCentsAsDollars(cents: int | None) -> str | None:
    if cents is None:
        return None
    return f"{cents / 100:.2f}"


# Cardinal / ordinal direction tokens → ERP Direction abbreviations.
_DIRECTION_ALIASES: dict[str, str] = {
    "n": "N",
    "s": "S",
    "e": "E",
    "w": "W",
    "ne": "NE",
    "nw": "NW",
    "se": "SE",
    "sw": "SW",
    "north": "N",
    "south": "S",
    "east": "E",
    "west": "W",
    "northeast": "NE",
    "northwest": "NW",
    "southeast": "SE",
    "southwest": "SW",
    "n.": "N",
    "s.": "S",
    "e.": "E",
    "w.": "W",
}

# Leftover after a direction token that is only a suffix means the direction
# word is the street name: "South St", "East Ave", "West Rd".
_STREET_SUFFIXES = {
    "st",
    "street",
    "ave",
    "avenue",
    "blvd",
    "boulevard",
    "dr",
    "drive",
    "rd",
    "road",
    "ln",
    "lane",
    "ct",
    "court",
    "cir",
    "circle",
    "way",
    "pl",
    "place",
    "ter",
    "terrace",
    "pkwy",
    "parkway",
    "hwy",
    "highway",
}

# Leading street number:
#   227, 123A, 12-14, 123 1/2
#   204 & 246, 882 884, Lots 2-4
# A second number joined by a space must be its own word (882 884),
# so 9600 54th does not swallow the 54.
_STREET_NUMBER_RE = re.compile(
    r"""^
    (?P<number>
        (?:
            [Ll]ots?\s+\d+(?:-\d+)?
          |
            \d+
            (?:-\d+)?
            (?:\s*&\s*\d+(?:-\d+)?)?
            (?:\s+\d+(?=\s))?
            (?:\s+\d+\/\d+)?
            [A-Za-z]?
        )
    )
    \s+
    (?P<rest>.+)$
    """,
    re.VERBOSE,
)


def _normalizeDirectionToken(token: str) -> str | None:
    key = token.strip().lower().rstrip(".")
    if key.endswith(".") and key[:-1] in _DIRECTION_ALIASES:
        key = key[:-1]
    return _DIRECTION_ALIASES.get(token.strip().lower()) or _DIRECTION_ALIASES.get(key)


def _isStreetSuffixOnly(tokens: list[str]) -> bool:
    if len(tokens) != 1:
        return False
    return tokens[0].lower().rstrip(".") in _STREET_SUFFIXES


def parseUsStreetAddress(
    address1: str | None, address2: str | None = None
) -> dict[str, str | None]:
    """Split Dwelling Blocks address1/address2 into ERP street columns.

    Dwelling Blocks only provides:
      - address1: full street line (e.g. "227 South Home Avenue", "522 N Washington St")
      - address2: secondary line, typically unit/apt (e.g. "307")

    Returns keys: streetNumber, direction, unit, streetAddress.
    """
    unit = (address2 or "").strip() or None
    raw = " ".join((address1 or "").split())
    if not raw:
        return {
            "streetNumber": None,
            "direction": None,
            "unit": unit,
            "streetAddress": None,
        }

    street_number: str | None = None
    direction: str | None = None
    street_address = raw

    match = _STREET_NUMBER_RE.match(raw)
    if match:
        street_number = match.group("number")
        remainder = match.group("rest").strip()
        tokens = remainder.split()
        if tokens:
            pre_dir = _normalizeDirectionToken(tokens[0])
            leftover = tokens[1:]
            if pre_dir and leftover and not _isStreetSuffixOnly(leftover):
                # Pre-directional: "South Home Avenue" / "N Washington St"
                direction = pre_dir
                street_address = " ".join(leftover).strip() or None
            else:
                # No pre-direction; check post-directional last token.
                post_dir = (
                    _normalizeDirectionToken(tokens[-1]) if len(tokens) > 1 else None
                )
                if post_dir:
                    direction = post_dir
                    street_address = " ".join(tokens[:-1]).strip() or None
                else:
                    street_address = remainder or None
        else:
            street_address = None
    else:
        # No leading number — keep full line as street name.
        tokens = raw.split()
        if len(tokens) > 1:
            pre_dir = _normalizeDirectionToken(tokens[0])
            leftover = tokens[1:]
            if pre_dir and leftover and not _isStreetSuffixOnly(leftover):
                direction = pre_dir
                street_address = " ".join(leftover).strip() or None

    return {
        "streetNumber": street_number,
        "direction": direction,
        "unit": unit,
        "streetAddress": street_address,
    }


def enrichAssignmentWithDetail(
    client: DwellingBlocksClient, item: dict[str, Any]
) -> dict[str, Any]:
    """Merge list item with detail: all Borrowers, access contacts, key properties.

    Dwelling Blocks has no CoBorrower ContactType. Additional borrowers are more
    contacts with contactType == "Borrower" on GET /vendors/orders/{assignmentId}.

    Total Payout Amount comes from Vendor transfers on
    GET /vendors/orders/{assignmentId}/transfers (amounts in cents).
    """
    assignment_id = item.get("assignmentId") or item.get("id")
    if assignment_id is None:
        raise ValueError(f"Assignment item missing id: {item}")

    detail = client.getVendorOrder(int(assignment_id))
    transfers = client.getVendorOrderTransfers(int(assignment_id))
    contacts = detail.get("contacts") or []
    # Primary borrower = highest contact id; co-borrower = next (descending).
    borrowers = contactsOfType(contacts, "Borrower", sort_by_id_desc=True)
    access_contacts = contactsOfType(contacts, "AccessContact", sort_by_id_desc=True)
    properties = detail.get("properties") or []

    primary = borrowers[0] if borrowers else None
    co_borrower = borrowers[1] if len(borrowers) > 1 else None
    access = access_contacts[0] if access_contacts else None
    payout_cents = vendorPayoutCents(transfers)
    address_parts = parseUsStreetAddress(
        detail.get("address1") if detail.get("address1") is not None else item.get("address1"),
        detail.get("address2") if detail.get("address2") is not None else item.get("address2"),
    )

    enriched = {
        **item,
        "detail": detail,
        "transfers": transfers,
        "contacts": contacts,
        "borrowers": borrowers,
        "accessContacts": access_contacts,
        "borrowerCount": len(borrowers),
        # Flat ERP-oriented fields (1st Borrower, 2nd Borrower as Co-Borrower)
        "borrowerName": (primary or {}).get("name"),
        "borrowerPhone": (primary or {}).get("phoneNumber"),
        "borrowerEmail": (primary or {}).get("email"),
        "coBorrowerName": (co_borrower or {}).get("name"),
        "coBorrowerPhone": (co_borrower or {}).get("phoneNumber"),
        "coBorrowerEmail": (co_borrower or {}).get("email"),
        "accessContactName": (access or {}).get("name"),
        "accessContactPhone": (access or {}).get("phoneNumber"),
        "accessContactEmail": (access or {}).get("email"),
        "loanType": propertyValueByLabel(properties, "Loan Type"),
        "propertyType": propertyValueByLabel(properties, "Property Type"),
        "appraisalPurpose": propertyValueByLabel(properties, "Appraisal Purpose"),
        "fhaCaseNumber": detail.get("fhaCaseNumber") or item.get("fhaNumber"),
        "orderSourceName": detail.get("orderSourceName"),
        "totalPayoutCents": payout_cents,
        "totalPayoutAmount": formatCentsAsDollars(payout_cents),
        "streetNumber": address_parts["streetNumber"],
        "direction": address_parts["direction"],
        "unit": address_parts["unit"],
        "streetAddress": address_parts["streetAddress"],
    }
    return enriched


def loadCredentials() -> tuple[str, str]:
    client_id = os.getenv("DWELLING_BLOCKS_CLIENT_ID", "").strip()
    client_secret = os.getenv("DWELLING_BLOCKS_CLIENT_SECRET", "").strip()
    if not client_id or not client_secret:
        raise SystemExit(
            "Missing DWELLING_BLOCKS_CLIENT_ID or DWELLING_BLOCKS_CLIENT_SECRET"
        )
    return client_id, client_secret


ERP_COLUMNS = (
    "ID",
    "File Name",
    "Assignee",
    "Report Type",
    "Loan Type",
    "Loan Number",
    "File Number",
    "Property Type",
    "Appraisal Purpose",
    "Client Reference Number",
    "FHA Case Number",
    "Total Payout Amount",
    "Additonal Information",
    "Street #",
    "Direction",
    "Unit",
    "Street Address",
    "City",
    "State",
    "Zip",
    "Country",
    "Borrower's Name",
    "Borrower's Phone Number",
    "Borrower's Email Address",
    "Co-Borrower's Name",
    "Co-Borrower's Phone Number",
    "Co-Borrower's Email Address",
    "Access Contact's Name",
    "Access Contact's Phone Number",
    "Access Contact's Email Address",
    "Order Date",
    "Report Uploaded",
    "Appointment Date",
    "Due Date",
)


def loadErpColumns() -> list[str]:
    """Return the ERP output headers in workbook order."""
    return list(ERP_COLUMNS)


def _cell(value: Any) -> str:
    """Normalize a value for ERP CSV/JSON cells; missing → blank."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    text = str(value).strip()
    return text


def _normalizeCountry(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.upper() in {"US", "USA", "UNITED STATES"}:
        return "USA"
    return text


def _buildErpFileName(
    street_number: Any, street_address: Any
) -> str | None:
    number = str(street_number or "").strip()
    address = str(street_address or "").strip()
    if number and address:
        return f"{number} {address}.pdf"
    if address:
        return f"{address}.pdf"
    return None


def assignmentToErpRow(
    item: dict[str, Any], columns: list[str] | None = None
) -> dict[str, str]:
    """Map an enriched assignment to ERP fields; unknowns stay blank."""
    columns = columns or loadErpColumns()

    assignee = item.get("assignedVendorName")
    if not assignee:
        detail_assignee = (item.get("detail") or {}).get("assignee") or item.get(
            "assignee"
        )
        if isinstance(detail_assignee, dict):
            assignee = detail_assignee.get("name")
        elif isinstance(detail_assignee, str):
            assignee = detail_assignee

    if item.get("streetNumber") is None and item.get("streetAddress") is None:
        address_parts = parseUsStreetAddress(
            item.get("address1"), item.get("address2")
        )
    else:
        address_parts = {
            "streetNumber": item.get("streetNumber"),
            "direction": item.get("direction"),
            "unit": item.get("unit"),
            "streetAddress": item.get("streetAddress"),
        }

    street_number = address_parts.get("streetNumber")
    street_address = address_parts.get("streetAddress") or item.get("address1")
    fha_case = item.get("fhaCaseNumber") or item.get("fhaNumber")

    values = {
        "ID": None,  # loan number + COMPANY_ID, assigned when appending
        "File Name": _buildErpFileName(street_number, street_address),
        "Assignee": assignee,
        "Report Type": item.get("reportType"),
        "Loan Type": item.get("loanType"),
        "Loan Number": item.get("loanNumber"),
        "File Number": None,  # intentionally blank
        "Property Type": item.get("propertyType"),
        "Appraisal Purpose": item.get("appraisalPurpose"),
        "Client Reference Number": None,  # intentionally blank
        "FHA Case Number": fha_case if fha_case else "Not set",
        "Total Payout Amount": item.get("totalPayoutAmount"),
        "Additonal Information": None,  # intentionally blank
        "Street #": street_number,
        "Direction": address_parts.get("direction"),
        "Unit": address_parts.get("unit"),
        "Street Address": street_address,
        "City": item.get("city"),
        "State": item.get("state"),
        "Zip": item.get("zipCode"),
        "Country": _normalizeCountry(item.get("country")),
        "Borrower's Name": item.get("borrowerName"),
        "Borrower's Phone Number": item.get("borrowerPhone"),
        "Borrower's Email Address": item.get("borrowerEmail"),
        "Co-Borrower's Name": item.get("coBorrowerName"),
        "Co-Borrower's Phone Number": item.get("coBorrowerPhone"),
        "Co-Borrower's Email Address": item.get("coBorrowerEmail"),
        "Access Contact's Name": item.get("accessContactName"),
        "Access Contact's Phone Number": item.get("accessContactPhone"),
        "Access Contact's Email Address": item.get("accessContactEmail"),
        "Order Date": item.get("createdOn"),
        "Report Uploaded": None,  # intentionally blank
        "Appointment Date": None,  # intentionally blank
        "Due Date": item.get("dueDate"),
    }

    return {column: _cell(values.get(column)) for column in columns}
