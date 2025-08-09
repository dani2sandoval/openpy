"""Simple SQLite database layer for the billing system."""
from __future__ import annotations

import sqlite3
from dataclasses import asdict
from typing import List, Tuple


class Database:
    """Handles persistence of billing data using SQLite."""

    def __init__(self, path: str = "billing.db") -> None:
        self.conn = sqlite3.connect(path)
        self._create_tables()

    def _create_tables(self) -> None:
        cur = self.conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                price REAL NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY,
                customer_id INTEGER NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS invoice_items (
                invoice_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL,
                FOREIGN KEY(invoice_id) REFERENCES invoices(id),
                FOREIGN KEY(product_id) REFERENCES products(id)
            )
            """
        )
        self.conn.commit()

    # Customer methods
    def add_customer(self, name: str, email: str) -> int:
        cur = self.conn.cursor()
        cur.execute("INSERT INTO customers(name, email) VALUES (?, ?)", (name, email))
        self.conn.commit()
        return cur.lastrowid

    # Product methods
    def add_product(self, name: str, price: float) -> int:
        cur = self.conn.cursor()
        cur.execute("INSERT INTO products(name, price) VALUES (?, ?)", (name, price))
        self.conn.commit()
        return cur.lastrowid

    # Invoice methods
    def create_invoice(self, customer_id: int) -> int:
        cur = self.conn.cursor()
        cur.execute("INSERT INTO invoices(customer_id) VALUES (?)", (customer_id,))
        self.conn.commit()
        return cur.lastrowid

    def add_invoice_item(self, invoice_id: int, product_id: int, quantity: int) -> None:
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO invoice_items(invoice_id, product_id, quantity) VALUES (?, ?, ?)",
            (invoice_id, product_id, quantity),
        )
        self.conn.commit()

    def list_invoices(self) -> List[Tuple]:
        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT invoices.id, customers.name, customers.email,
                   products.name, invoice_items.quantity, products.price
            FROM invoices
            JOIN customers ON invoices.customer_id = customers.id
            JOIN invoice_items ON invoice_items.invoice_id = invoices.id
            JOIN products ON invoice_items.product_id = products.id
            ORDER BY invoices.id
            """
        )
        return cur.fetchall()

    def close(self) -> None:
        self.conn.close()
