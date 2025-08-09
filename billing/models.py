from dataclasses import dataclass, field
from typing import List


@dataclass
class Product:
    """Represents a product that can be sold."""

    id: int
    name: str
    price: float


@dataclass
class Customer:
    """Represents a customer."""

    id: int
    name: str
    email: str


@dataclass
class InvoiceItem:
    """Represents an item within an invoice."""

    product: Product
    quantity: int

    @property
    def total(self) -> float:
        """Total cost for this invoice item."""
        return self.product.price * self.quantity


@dataclass
class Invoice:
    """Represents an invoice containing multiple items."""

    id: int
    customer: Customer
    items: List[InvoiceItem] = field(default_factory=list)

    def add_item(self, product: Product, quantity: int) -> None:
        """Add a product to the invoice."""
        self.items.append(InvoiceItem(product, quantity))

    @property
    def total(self) -> float:
        """Compute the total amount of the invoice."""
        return sum(item.total for item in self.items)
