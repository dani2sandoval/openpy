from billing.models import Customer, Product, Invoice


def test_invoice_total():
    customer = Customer(id=1, name="Alice", email="alice@example.com")
    product = Product(id=1, name="Widget", price=10.0)
    invoice = Invoice(id=1, customer=customer)
    invoice.add_item(product, 2)

    assert invoice.total == 20.0
