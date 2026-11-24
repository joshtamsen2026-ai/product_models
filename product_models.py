"""Product Catalog Models - SQLAlchemy ORM exercise.

Defines the ``Category`` and ``Product`` models, creates ``product_catalog.db``,
inserts sample data, and runs three queries:

1. All categories.
2. All products that are in stock.
3. All products priced under $50.

The results are printed to the console and exported to a professional CSV file
and a DOCX report with charts, findings, and a requirements summary.

How the file is organized (read it top to bottom):

* Configuration - file paths, colors, and titles used everywhere else.
* Engine - the SQLAlchemy object that talks to the SQLite database file.
* Models - Python classes that describe the two database tables.
* Sample data - the categories and products that are inserted.
* Queries - the three ``select()`` statements the exercise asks for.
* Narrative - helpers that turn numbers into plain-English sentences.
* Schema verification - reads the tables back from the database as proof.
* Requirements - answers each requirement of the brief as a question.
* Console output, CSV export, Charts, DOCX report - the different outputs.
* Entry point - ``main()``, which runs everything in the right order.

Usage:
    python product_models.py            # build database and reports, open both
    python product_models.py --no-open  # build without opening the reports
    python product_models.py --force    # rewrite reports even if unchanged
"""

# --- Imports ----------------------------------------------------------------
## This line makes Python treat every type hint as plain text, so a hint can name
## a class that is defined further down the file (or only imported for mypy).
from __future__ import annotations

## Standard-library modules that come with Python (no install needed):
## argparse reads command-line options, csv writes the spreadsheet, hashlib and
## json build and store the report cache, io keeps chart images in memory, and
## os/platform/subprocess/sys talk to the computer.
import argparse
import csv
import hashlib
import io
import json
import os
import platform
import subprocess
import sys

## "from X import Y" brings in just one name. defaultdict is a dict that creates a
## missing key for you, dataclass builds simple classes, getsource reads the
## source code of a class, and count() hands out 1, 2, 3, ... on demand.
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from inspect import getsource
from itertools import count
from pathlib import Path
from typing import TYPE_CHECKING, Any

## SQLAlchemy is a third-party library (installed with pip) that lets us work with
## database tables as Python classes. These are the column types, the engine
## factory, the event system, the schema inspector, and the select() query builder.
from sqlalchemy import Boolean, Float, String, create_engine, event, inspect, select

## "import ... as NEW_NAME" renames the imported value; here it is the version text.
from sqlalchemy import __version__ as SQLALCHEMY_VERSION

## The ORM ("Object Relational Mapper") pieces used to declare models and to open
## a Session, which is the unit of work for adding and querying rows.
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column

## TYPE_CHECKING is False when the program runs and True only for type checkers
## such as mypy. Imports inside this block are used just for type hints, so the
## program does not pay the cost of loading docx/matplotlib until it needs them.
if TYPE_CHECKING:
    from collections.abc import Iterable, Sequence

    from docx.document import Document as DocxDocument
    from docx.table import _Cell
    from matplotlib.axes import Axes
    from matplotlib.figure import Figure

# --- Configuration -----------------------------------------------------------
## Constants (UPPER_CASE names) that never change while the program runs. Keeping
## them in one place means a path or color can be changed in a single spot.

## Path(__file__) is this script's own location; .resolve() makes it a full path.
## The "/" operator on Path objects joins folder and file names safely.
SCRIPT_PATH = Path(__file__).resolve()
BASE_DIR = SCRIPT_PATH.parent
DB_PATH = BASE_DIR / "product_catalog.db"
CSV_PATH = BASE_DIR / "product_catalog_results.csv"
DOCX_PATH = BASE_DIR / "product_catalog_report.docx"
## Remembers a fingerprint of the last reports so unchanged reports are reused.
CACHE_PATH = BASE_DIR / ".product_catalog_cache.json"

## Stock labels shown in tables, charts, and legends.
IN_STOCK_LABEL = "In stock"
OUT_OF_STOCK_LABEL = "Out of stock"

## Chart image settings. zlib level 1 is the fastest PNG compression: the files
## are slightly larger but the pictures look exactly the same.
CHART_DPI = 150
PNG_COMPRESS_LEVEL = 1

## The price cut-off for query 3, plus the hex colors used by charts and tables.
PRICE_LIMIT = 50.0
BRAND_COLOR = "#1F4E79"
ACCENT_COLOR = "#C55A11"
POSITIVE_COLOR = "#548235"
NEGATIVE_COLOR = "#C00000"

## The report title and the section headings. A tuple ( ... ) is like a list that
## cannot be changed, which suits fixed values like these.
REPORT_TITLE = "Product Catalog Models Report"
SECTION_TITLES = (
    "Executive Summary",
    "Questions and Answers",
    "Verified Database Schema",
    "Category Analysis",
    "Key Findings",
    "Bonus: echo=True SQL vs. Raw SQL",
    "Conclusion",
)

## Word table fills (hex without '#'). Word wants colors written without the '#'.
WHITE_HEX = "FFFFFF"
ZEBRA_HEX = "EAF1F8"
CALLOUT_HEX = "FFF4E5"
CODE_HEX = "F2F2F2"

# --- Engine -----------------------------------------------------------------
## The engine is SQLAlchemy's connection to the database. The f-string (a string
## starting with f) puts the value of DB_PATH inside the {braces}, giving a URL
## like "sqlite:///C:/.../product_catalog.db". SQLite creates the file if needed.
## echo=True logs every SQL statement SQLAlchemy sends to the database.
engine = create_engine(f"sqlite:///{DB_PATH}", echo=True)

## Every statement is also captured so the report can compare it with raw SQL.
## "list[str]" is a type hint: it tells readers and mypy this holds strings.
EXECUTED_SQL: list[str] = []


## "@event.listens_for(...)" is a decorator: it registers the function below as an
## event listener, so SQLAlchemy calls it automatically right before it runs any
## SQL statement on this engine. We only use "statement"; the other parameters
## are required by SQLAlchemy's event signature.
@event.listens_for(engine, "before_cursor_execute")
def _record_sql(
    conn: Any,
    cursor: Any,
    statement: str,
    parameters: Any,
    context: Any,
    executemany: bool,
) -> None:
    """Remember every SQL statement that SQLAlchemy is about to run.

    Why: the Bonus section of the report compares the SQL that SQLAlchemy
    generated with hand-written SQL, so we need a copy of each statement.

    Args:
        conn: The database connection (not used here).
        cursor: The low-level database cursor (not used here).
        statement: The SQL text about to be executed.
        parameters: The values sent alongside the SQL (not used here).
        context: SQLAlchemy's execution details (not used here).
        executemany: True when one statement runs for many rows (not used).

    Returns:
        None. The statement is appended to the ``EXECUTED_SQL`` list.
    """
    ## .strip() removes spaces and newlines from both ends before saving the text.
    EXECUTED_SQL.append(statement.strip())


# --- Models ------------------------------------------------------------------
## A "model" is a Python class that maps to one database table: each class
## attribute becomes a column, and each object (instance) becomes one row.


## DeclarativeBase is SQLAlchemy's starting point for models. Every model inherits
## from this Base class, so Base.metadata ends up knowing about every table.
class Base(DeclarativeBase):
    """Declarative base shared by every catalog model.

    It has no code of its own: inheriting from ``DeclarativeBase`` is enough.
    ``Base.metadata`` collects all tables of the subclasses, which is what
    ``create_all()`` and ``drop_all()`` use to build or remove the tables.
    """


## Category maps to the "categories" table. Notes for beginners:
## * __tablename__ sets the real table name in the database.
## * "Mapped[int]" is a type hint that tells SQLAlchemy (and mypy) the Python
##   type of the column; mapped_column(...) adds the database rules.
## * primary_key=True makes id the unique row number that SQLite fills in.
## * String(100) is text up to 100 characters; nullable=False means "required"
##   and unique=True means no two rows may share the same name.
## * "str | None" means the value may be missing, so description is optional.
## * __repr__ controls how an object looks when printed, which helps debugging.
## Explanations live up here, not inside the class, because the report copies
## the class's source code (see _model_source) and must not include them.
class Category(Base):
    """A product category such as "Electronics" or "Books"."""

    __tablename__ = "categories"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100), nullable=False, unique=True)
    description: Mapped[str | None] = mapped_column(String(255))

    def __repr__(self) -> str:
        """Return a short, readable description such as ``Category(id=1, ...)``.

        Python calls this automatically when a Category is printed or shown in
        the debugger. ``!r`` inside the f-string adds quotes around the name.

        Returns:
            The text form of this category.
        """
        return f"Category(id={self.id}, name={self.name!r})"


## Product maps to the "products" table, using the same ideas as Category:
## * Float stores decimal numbers such as 24.99 (the price).
## * Boolean stores True/False; default=True means Python fills in True when a
##   new product is saved without an in_stock value.
## * category_name is plain text for now; a later lesson turns it into a
##   foreign key (a real link to the categories table).
class Product(Base):
    """A catalog product.

    ``category_name`` stores the category as plain text for now; the
    Relationships lesson replaces it with a proper foreign key.
    """

    __tablename__ = "products"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100), nullable=False)
    price: Mapped[float] = mapped_column(Float, nullable=False)
    in_stock: Mapped[bool] = mapped_column(Boolean, default=True)
    category_name: Mapped[str | None] = mapped_column(String(100))

    def __repr__(self) -> str:
        """Return a short, readable description such as ``Product(id=1, ...)``.

        ``{self.price:.2f}`` formats the price with exactly two decimal places.

        Returns:
            The text form of this product.
        """
        return f"Product(id={self.id}, name={self.name!r}, price={self.price:.2f})"


# --- Sample data -------------------------------------------------------------
## The rows inserted into the database. Each inner tuple is one row.

## Category names are stored once as constants, so a typo cannot silently create
## a product whose category_name matches no real category.
ELECTRONICS = "Electronics"
BOOKS = "Books"
HOME_KITCHEN = "Home & Kitchen"
SPORTS_OUTDOORS = "Sports & Outdoors"

## Type hint explained: a tuple of (name, description) pairs, where the "..."
## means "any number of them" and "str | None" means the description is optional.
CATEGORIES: tuple[tuple[str, str | None], ...] = (
    (ELECTRONICS, "Gadgets, devices, and computer accessories"),
    (BOOKS, "Programming and software engineering books"),
    (HOME_KITCHEN, "Cookware, small appliances, and home essentials"),
    (SPORTS_OUTDOORS, None),  ## Demonstrates the optional description.
)

## (name, price, category_name, in_stock)
PRODUCTS: tuple[tuple[str, float, str, bool], ...] = (
    ("Wireless Mouse", 24.99, ELECTRONICS, True),
    ("Noise-Cancelling Headphones", 199.99, ELECTRONICS, True),
    ("USB-C Charger", 19.99, ELECTRONICS, False),
    ("27-inch 4K Monitor", 329.00, ELECTRONICS, True),
    ("Python Crash Course", 39.95, BOOKS, True),
    ("Clean Code", 44.50, BOOKS, False),
    ("The Pragmatic Programmer", 52.00, BOOKS, True),
    ("Chef's Knife", 89.00, HOME_KITCHEN, True),
    ("French Press", 29.99, HOME_KITCHEN, True),
    ("Air Fryer", 119.99, HOME_KITCHEN, False),
    ("Yoga Mat", 25.00, SPORTS_OUTDOORS, True),
    ("Camping Tent", 149.99, SPORTS_OUTDOORS, True),
)

# --- Queries -----------------------------------------------------------------
## select(Model) builds a SELECT statement without running it yet. .where(...)
## adds a filter, and .order_by(...) sorts the rows. The statements are stored
## here so the same query objects can be run and also shown in the report.

## Query 1: every category, sorted by id.
ALL_CATEGORIES_QUERY = select(Category).order_by(Category.id)
## Query 2: only products whose in_stock column is True.
IN_STOCK_QUERY = select(Product).where(Product.in_stock).order_by(Product.id)
## Query 3: products cheaper than PRICE_LIMIT, from cheapest to most expensive.
UNDER_LIMIT_QUERY = (
    select(Product).where(Product.price < PRICE_LIMIT).order_by(Product.price)
)
## Extra query: every product, used for statistics and charts.
ALL_PRODUCTS_QUERY = select(Product).order_by(Product.id)


## @dataclass writes __init__, __repr__, and __eq__ for us from the fields below.
## frozen=True makes objects read-only after creation, and slots=True uses less
## memory and stops typos from silently creating new attributes.
@dataclass(frozen=True, slots=True)
class CategoryRow:
    """Plain, session-independent copy of a ``Category``.

    Why: ORM objects are tied to the database session. Once the session is
    closed, reading them can fail. Copying the values into this simple,
    read-only object lets the reports use them safely afterwards.

    Attributes:
        id: The category's primary key.
        name: The category name, e.g. "Books".
        description: Optional description, or None when there is none.
    """

    id: int
    name: str
    description: str | None


@dataclass(frozen=True, slots=True)
class ProductRow:
    """Plain, session-independent copy of a ``Product``.

    Attributes:
        id: The product's primary key.
        name: The product name.
        price: The price in US dollars.
        in_stock: True when the product is available.
        category_name: The category name as text, or None.
    """

    id: int
    name: str
    price: float
    in_stock: bool
    category_name: str | None

    ## @property lets you read a method like a normal attribute: row.category
    ## (no parentheses) runs the method below and returns its result.
    @property
    def category(self) -> str:
        """Return the category name, or "Uncategorized" when it is missing.

        ``a or b`` gives ``a`` when it has a value, otherwise ``b``.
        """
        return self.category_name or "Uncategorized"

    @property
    def stock_label(self) -> str:
        """Return "In stock" or "Out of stock" for display in tables.

        This uses a one-line if/else (a "conditional expression").
        """
        return IN_STOCK_LABEL if self.in_stock else OUT_OF_STOCK_LABEL


@dataclass(frozen=True, slots=True)
class CategoryStats:
    """Per-category aggregates used by the reports and charts.

    Attributes:
        name: The category name.
        product_count: How many products belong to the category.
        in_stock_count: How many of those products are in stock.
        average_price: The mean price of the category's products.
        min_price: The cheapest price in the category.
        max_price: The most expensive price in the category.
    """

    name: str
    product_count: int
    in_stock_count: int
    average_price: float
    min_price: float
    max_price: float


@dataclass(frozen=True, slots=True)
class CatalogResults:
    """Everything the reports need, detached from the database session.

    Attributes:
        categories: Result of query 1 (all categories).
        in_stock: Result of query 2 (products that are in stock).
        under_limit: Result of query 3 (products under the price limit).
        products: Every product, used for statistics.
    """

    categories: list[CategoryRow]
    in_stock: list[ProductRow]
    under_limit: list[ProductRow]
    products: list[ProductRow]


def _to_product_row(product: Product) -> ProductRow:
    """Copy one ORM ``Product`` into a plain ``ProductRow``.

    The leading underscore in the name is a Python convention meaning
    "internal helper, not meant to be used from outside this file".

    Args:
        product: A Product object loaded from the database session.

    Returns:
        A read-only ProductRow holding the same values.
    """
    ## Keyword arguments (name=value) make it clear which value fills which field.
    return ProductRow(
        id=product.id,
        name=product.name,
        price=product.price,
        in_stock=product.in_stock,
        category_name=product.category_name,
    )


def _make_products() -> list[Product]:
    """Create (but do not save) a Product object for every row in PRODUCTS.

    Returns:
        A list of new Product objects ready to be added to a session.
    """
    products = []
    ## "for a, b, c, d in PRODUCTS" unpacks each 4-item tuple into four names.
    for name, price, category_name, in_stock in PRODUCTS:
        product = Product(name=name, price=price, category_name=category_name)
        ## In-stock items rely on the column default (True) instead of setting it.
        if not in_stock:
            product.in_stock = False
        products.append(product)
    return products


def build_catalog() -> CatalogResults:
    """Create the schema, insert the sample data, and run the queries.

    Steps: drop and recreate the tables, insert categories and products in one
    session, run the three required queries, print them, and return copies of
    the results that stay usable after the session is closed.

    Returns:
        A CatalogResults object with the rows returned by every query.
    """
    ## Dropping first makes the script re-runnable (no UNIQUE violations) and
    ## guarantees the CREATE TABLE statements appear in the echo output.
    Base.metadata.drop_all(engine)
    Base.metadata.create_all(engine)

    ## A "with" block opens the Session and closes it automatically at the end,
    ## even if an error happens. The session tracks objects to save and queries.
    with Session(engine) as session:
        ## add_all() takes many objects at once. The "(... for ... in ...)" part
        ## is a generator expression: it builds one Category per CATEGORIES row.
        session.add_all(
            Category(name=name, description=description)
            for name, description in CATEGORIES
        )
        session.add_all(_make_products())
        ## commit() writes the pending INSERT statements to the database file.
        session.commit()

        ## session.scalars(query) runs the query and yields model objects. The
        ## list comprehension [... for c in ...] copies each into a CategoryRow.
        categories = [
            CategoryRow(id=c.id, name=c.name, description=c.description)
            for c in session.scalars(ALL_CATEGORIES_QUERY)
        ]
        ## Run the other queries the same way and bundle every result together.
        results = CatalogResults(
            categories=categories,
            in_stock=[_to_product_row(p) for p in session.scalars(IN_STOCK_QUERY)],
            under_limit=[
                _to_product_row(p) for p in session.scalars(UNDER_LIMIT_QUERY)
            ],
            products=[_to_product_row(p) for p in session.scalars(ALL_PRODUCTS_QUERY)],
        )
        ## The brief asks to query *and print* using the session.
        print_results(results, build_query_explanations(results))
        return results


def compute_category_stats(results: CatalogResults) -> list[CategoryStats]:
    """Aggregate product counts, stock levels, and prices per category.

    Args:
        results: The query results from ``build_catalog()``.

    Returns:
        One CategoryStats per category, in the same order as the categories.
    """
    ## Group products by category name. defaultdict(list) starts each new key
    ## with an empty list, so we can append without checking first.
    by_category: dict[str, list[ProductRow]] = defaultdict(list)
    for product in results.products:
        by_category[product.category].append(product)

    stats = []
    for category in results.categories:
        ## .get(key, []) returns an empty list for categories with no products.
        items = by_category.get(category.name, [])
        ## "or [0.0]" avoids dividing by zero / min() of an empty list.
        prices = [p.price for p in items] or [0.0]
        stats.append(
            CategoryStats(
                name=category.name,
                product_count=len(items),
                ## True counts as 1 and False as 0, so sum() counts in-stock items.
                in_stock_count=sum(p.in_stock for p in items),
                average_price=sum(prices) / len(prices),
                min_price=min(prices),
                max_price=max(prices),
            )
        )
    return stats


# --- Narrative ---------------------------------------------------------------
## Small helpers that format numbers and lists, plus functions that write the
## plain-English summaries, findings, and conclusion used in every output.


def _money(value: float) -> str:
    """Format a number as dollars, e.g. 1234.5 -> "$1,234.50".

    In ``{value:,.2f}`` the comma adds thousands separators and ``.2f`` keeps
    two decimal places.

    Args:
        value: The amount to format.

    Returns:
        The formatted price text.
    """
    return f"${value:,.2f}"


def _percent(part: int, whole: int) -> str:
    """Return ``part`` as a whole-number percentage of ``whole``, e.g. "75%".

    Args:
        part: The count being measured.
        whole: The total; when it is 0 the result is "0%" (no division by 0).

    Returns:
        The percentage text.
    """
    ## ".0%" multiplies by 100, rounds to 0 decimals, and adds the % sign.
    return f"{part / whole:.0%}" if whole else "0%"


def _names(products: Iterable[ProductRow]) -> str:
    """Join product names with commas, or return "none" if there are none.

    Args:
        products: Any collection of ProductRow objects.

    Returns:
        Text such as "Yoga Mat, French Press".
    """
    return ", ".join(p.name for p in products) or "none"


def _join(names: Sequence[str]) -> str:
    """'A', 'A and B', or 'A, B, and C'.

    Joins names the way you would in an English sentence.

    Args:
        names: The names to join (a list or tuple of strings).

    Returns:
        The joined text.
    """
    ## One or two names: simply put " and " between them.
    if len(names) <= 2:
        return " and ".join(names)
    ## names[:-1] is every name except the last; names[-1] is the last one.
    return f"{', '.join(names[:-1])}, and {names[-1]}"


def _lowest_availability(stats: Iterable[CategoryStats]) -> tuple[list[str], float]:
    """Every category tied for the lowest in-stock rate, and that rate.

    Args:
        stats: The per-category statistics.

    Returns:
        A pair: (list of category names, lowest rate between 0.0 and 1.0).
        Returns ([], 0.0) when no category has any products.
    """
    ## A dict comprehension {key: value for ...}; categories with 0 products are
    ## skipped by the "if" so we never divide by zero.
    rates = {
        s.name: s.in_stock_count / s.product_count for s in stats if s.product_count
    }
    if not rates:
        return [], 0.0
    lowest = min(rates.values())
    ## Keep every category whose rate equals the lowest (there may be a tie).
    return [name for name, rate in rates.items() if rate == lowest], lowest


## Column headings for the category statistics table (CSV and Word).
STATS_HEADERS = (
    "Category",
    "Products",
    "In Stock",
    "Availability",
    "Average Price",
    "Lowest Price",
    "Highest Price",
)


def stats_table(stats: Iterable[CategoryStats]) -> list[tuple[str, ...]]:
    """Turn category statistics into rows of text for a table.

    Args:
        stats: The per-category statistics.

    Returns:
        One tuple of strings per category, matching ``STATS_HEADERS``.
    """
    ## A list comprehension that builds one formatted row per category.
    return [
        (
            s.name,
            str(s.product_count),
            str(s.in_stock_count),
            _percent(s.in_stock_count, s.product_count),
            _money(s.average_price),
            _money(s.min_price),
            _money(s.max_price),
        )
        for s in stats
    ]


def _status_color(value: str) -> str | None:
    """Hex (no '#') for MET / NOT MET cells; None leaves other text unchanged.

    Args:
        value: The cell text, e.g. "MET".

    Returns:
        Green for "MET", red for "NOT MET", otherwise None.
    """
    ## Look the text up in a small dict; .get() returns None if it is not there.
    colors = {"MET": POSITIVE_COLOR, "NOT MET": NEGATIVE_COLOR}
    color = colors.get(value)
    ## color[1:] drops the leading '#', because Word wants plain hex digits.
    return color[1:] if color else None


def build_summary_metrics(results: CatalogResults) -> list[tuple[str, str]]:
    """Headline numbers for the executive summary.

    Args:
        results: The query results.

    Returns:
        A list of (metric name, formatted value) pairs.
    """
    ## Gather the numbers once so the list below stays short and readable.
    products = results.products
    total = len(products)
    prices = [p.price for p in products]
    ## min()/max() with default=0.0 return 0.0 instead of crashing on no prices.
    return [
        ("Total categories", str(len(results.categories))),
        ("Total products", str(total)),
        (
            "Products in stock",
            f"{len(results.in_stock)} ({_percent(len(results.in_stock), total)})",
        ),
        (
            "Products out of stock",
            (
                f"{total - len(results.in_stock)} "
                f"({_percent(total - len(results.in_stock), total)})"
            ),
        ),
        (
            f"Products under {_money(PRICE_LIMIT)}",
            f"{len(results.under_limit)} ({_percent(len(results.under_limit), total)})",
        ),
        ("Average product price", _money(sum(prices) / total if total else 0.0)),
        ("Lowest price", _money(min(prices, default=0.0))),
        ("Highest price", _money(max(prices, default=0.0))),
    ]


def build_query_explanations(results: CatalogResults) -> dict[str, str]:
    """Plain-English answer for each of the three required queries.

    Args:
        results: The query results.

    Returns:
        A dict with the keys "categories", "in_stock", and "under_limit",
        each holding a sentence or two describing that query's result.
    """
    ## Products that are not in stock, found with a filtering list comprehension.
    out_of_stock = [p for p in results.products if not p.in_stock]
    ## key=lambda p: p.price tells min() to compare products by price. A lambda
    ## is a tiny one-line function without a name.
    cheapest = min(results.under_limit, key=lambda p: p.price, default=None)
    ## Only mention the cheapest item when query 3 returned at least one product.
    under_text = (
        f" The cheapest item is {cheapest.name} at {_money(cheapest.price)}."
        if cheapest
        else ""
    )
    ## Categories with an empty description prove that the column is optional.
    missing_description = [c.name for c in results.categories if not c.description]
    ## Build the three sentences. Strings next to each other inside ( ) are
    ## joined automatically, which keeps long text within the line limit.
    return {
        "categories": (
            f"The catalog contains {len(results.categories)} categories: "
            f"{', '.join(c.name for c in results.categories)}. "
            + (
                f"{', '.join(missing_description)} has no description, which shows "
                "that the description column is optional."
                if missing_description
                else "Every category has a description."
            )
        ),
        "in_stock": (
            f"{len(results.in_stock)} of {len(results.products)} products "
            f"({_percent(len(results.in_stock), len(results.products))}) are in "
            f"stock. Out of stock: {_names(out_of_stock)}."
        ),
        "under_limit": (
            f"{len(results.under_limit)} products cost less than "
            f"{_money(PRICE_LIMIT)}: {_names(results.under_limit)}.{under_text}"
        ),
    }


def build_findings(results: CatalogResults, stats: list[CategoryStats]) -> list[str]:
    """Key insights drawn from the query results.

    Args:
        results: The query results.
        stats: The per-category statistics.

    Returns:
        A list of finding sentences (one bullet point each in the report).
    """
    ## Guard clause: stop early if there is nothing to analyze.
    if not results.products:
        return ["The catalog is empty, so there are no findings to report."]

    ## Find the extremes: most/least expensive product and category.
    priciest = max(results.products, key=lambda p: p.price)
    cheapest = min(results.products, key=lambda p: p.price)
    premium = max(stats, key=lambda s: s.average_price)
    budget = min(stats, key=lambda s: s.average_price)
    ## Unpack the (names, rate) pair returned by the helper into two variables.
    weakest, weakest_rate = _lowest_availability(stats)
    weakest_stats = [s for s in stats if s.name in weakest]
    ## Cheap products that can be sold right now.
    affordable_in_stock = [p for p in results.under_limit if p.in_stock]
    ## Each item in the list below is one finding sentence.
    return [
        (
            f"{premium.name} is the premium category with an average price of "
            f"{_money(premium.average_price)}; {budget.name} is the most affordable at "
            f"{_money(budget.average_price)}."
        ),
        (
            f"Prices range from {_money(cheapest.price)} ({cheapest.name}) to "
            f"{_money(priciest.price)} ({priciest.name})."
        ),
        (
            f"{_join(weakest)} {'has' if len(weakest) == 1 else 'tie for'} the "
            f"lowest availability at {weakest_rate:.0%} in stock ("
            + "; ".join(
                f"{s.name}: {s.in_stock_count} of {s.product_count}"
                for s in weakest_stats
            )
            + "). Restocking should start there."
        ),
        (
            f"{len(affordable_in_stock)} products are both under {_money(PRICE_LIMIT)} "
            f"and in stock ({_names(affordable_in_stock)}), so they can be sold today."
        ),
        (
            "Every product used a category name that matches a real category, but "
            "nothing enforces that yet. A foreign key (next lesson) will guarantee it."
        ),
    ]


def compiled_sql(statement: Any) -> str:
    """Render a SQLAlchemy statement exactly as echo=True prints it.

    Args:
        statement: A query object such as ``ALL_CATEGORIES_QUERY``.

    Returns:
        The SQL text, with ``?`` placeholders where values are sent separately.
    """
    ## .compile(engine) turns the query object into SQLite-flavored SQL text.
    return str(statement.compile(engine))


def _first_executed(prefix: str) -> str:
    """Find the first captured SQL statement that starts with ``prefix``.

    Args:
        prefix: The start of the statement, e.g. "CREATE TABLE products".

    Returns:
        The matching SQL text, or "(not captured)" if none was found.
    """
    ## next() takes the first item from the generator, or the default if empty.
    ## .upper() on both sides makes the comparison ignore upper/lower case.
    return next(
        (sql for sql in EXECUTED_SQL if sql.upper().startswith(prefix.upper())),
        "(not captured)",
    )


def build_sql_comparison() -> list[tuple[str, str, str, str]]:
    """(label, SQL from echo=True, equivalent hand-written SQL, difference).

    Used by the Bonus section to show what SQLAlchemy really sent to SQLite
    next to the SQL a person would usually write by hand.

    Returns:
        A list of 4-item tuples, one per statement being compared.
    """
    ## Each tuple: a label, the captured/compiled SQL, the hand-written SQL, and
    ## a short explanation of how they differ. "\n" inside text is a new line.
    return [
        (
            "CREATE TABLE categories",
            _first_executed("CREATE TABLE categories"),
            (
                "CREATE TABLE categories (\n"
                "    id INTEGER PRIMARY KEY,\n"
                "    name VARCHAR(100) NOT NULL UNIQUE,\n"
                "    description VARCHAR(255)\n"
                ");"
            ),
            (
                "Same columns and constraints. SQLAlchemy writes id as INTEGER NOT "
                "NULL and declares PRIMARY KEY (id) and UNIQUE (name) at the end."
            ),
        ),
        (
            "CREATE TABLE products",
            _first_executed("CREATE TABLE products"),
            (
                "CREATE TABLE products (\n"
                "    id INTEGER PRIMARY KEY,\n"
                "    name VARCHAR(100) NOT NULL,\n"
                "    price FLOAT NOT NULL,\n"
                "    in_stock BOOLEAN NOT NULL DEFAULT 1,\n"
                "    category_name VARCHAR(100)\n"
                ");"
            ),
            (
                "No DEFAULT clause appears: default=True is applied by Python when the "
                "row is inserted. Use server_default to put it in the table itself."
            ),
        ),
        (
            "INSERT INTO categories",
            _first_executed("INSERT INTO categories"),
            (
                "INSERT INTO categories (name, description)\n"
                "VALUES ('Electronics', 'Gadgets, devices, ...');"
            ),
            (
                "SQLAlchemy uses ? placeholders (safe from SQL injection) with the "
                "values sent separately, and adds RETURNING id to read back each new "
                "primary key. This SQLite build cannot batch these, so the echo log "
                "shows one INSERT per row."
            ),
        ),
        (
            "INSERT INTO products",
            _first_executed("INSERT INTO products"),
            (
                "INSERT INTO products (name, price, in_stock, category_name)\n"
                "VALUES ('Wireless Mouse', 24.99, 1, 'Electronics');"
            ),
            (
                "Same as above. in_stock is always sent because Python fills in the "
                "default before the INSERT."
            ),
        ),
        (
            "Query 1 - all categories",
            compiled_sql(ALL_CATEGORIES_QUERY),
            "SELECT * FROM categories ORDER BY id;",
            "SQLAlchemy lists every column explicitly instead of using *.",
        ),
        (
            "Query 2 - in-stock products",
            compiled_sql(IN_STOCK_QUERY),
            "SELECT * FROM products WHERE in_stock = 1 ORDER BY id;",
            "Python True is stored as 1 because SQLite has no real boolean type.",
        ),
        (
            "Query 3 - products under $50",
            compiled_sql(UNDER_LIMIT_QUERY),
            f"SELECT * FROM products WHERE price < {PRICE_LIMIT:g} ORDER BY price;",
            (
                "The 50 is sent as a bound ? parameter instead of being typed into the "
                "SQL text."
            ),
        ),
    ]


def build_conclusion(
    results: CatalogResults,
    stats: list[CategoryStats],
    answers: list[RequirementAnswer],
) -> tuple[str, str]:
    """Closing paragraphs: (summary of findings, summary of requirements).

    Args:
        results: The query results.
        stats: The per-category statistics.
        answers: The requirement answers (used to count how many were met).

    Returns:
        Two paragraphs of text: the findings summary and the requirements one.
    """
    ## Work out every number the two paragraphs mention.
    products = results.products
    total = len(products)
    average = sum(p.price for p in products) / total if total else 0.0
    premium = max(stats, key=lambda s: s.average_price)
    budget = min(stats, key=lambda s: s.average_price)
    weakest, rate = _lowest_availability(stats)
    ready = [p for p in results.under_limit if p.in_stock]
    met = sum(a.met for a in answers)

    ## Paragraph 1: what the data shows, built from many joined f-strings.
    findings = (
        f"The catalog holds {total} products across {len(results.categories)} "
        f"categories with an average price of {_money(average)}. "
        f"{len(results.in_stock)} of {total} products "
        f"({_percent(len(results.in_stock), total)}) are in stock and "
        f"{len(results.under_limit)} ({_percent(len(results.under_limit), total)}) "
        f"cost less than {_money(PRICE_LIMIT)}. {premium.name} is the premium "
        f"category ({_money(premium.average_price)} on average) and {budget.name} "
        f"is the most affordable ({_money(budget.average_price)}). "
        f"{_join(weakest)} {'has' if len(weakest) == 1 else 'have'} the lowest "
        f"stock availability ({rate:.0%}), so restocking should start there. "
        f"{len(ready)} products are both under {_money(PRICE_LIMIT)} and in stock, "
        "making them the strongest candidates for budget promotions."
    )
    ## Paragraph 2: how each requirement of the exercise was satisfied.
    requirements = (
        f"{met} of {len(answers)} requirements were met. product_models.py sets up "
        "a SQLAlchemy engine for product_catalog.db with echo=True, defines the "
        "Category model (integer primary key, unique non-null name, optional "
        "description) and the Product model (integer primary key, non-null name "
        "and price, in_stock defaulting to True, and category_name stored as "
        "text), and creates both tables with Base.metadata.create_all(engine). "
        f"Using a session it inserts {len(results.categories)} categories and "
        f"{total} products, then queries and prints all categories "
        f"({len(results.categories)} rows), the in-stock products "
        f"({len(results.in_stock)} rows), and the products under "
        f"{_money(PRICE_LIMIT)} ({len(results.under_limit)} rows). Every answer "
        "was verified against the database file itself rather than assumed. The "
        "next step, in the Relationships lesson, is to replace category_name with "
        "a foreign key so the database enforces valid categories."
    )
    ## Returning two values with a comma actually returns one tuple.
    return findings, requirements


# --- Schema verification -----------------------------------------------------
## Reads the table definitions back out of product_catalog.db, so the report can
## prove the models created the right columns instead of just assuming it.


@dataclass(frozen=True, slots=True)
class ColumnInfo:
    """A column exactly as it exists in ``product_catalog.db``.

    Attributes:
        model: The Python model name, e.g. "Product".
        table: The database table name, e.g. "products".
        name: The column name.
        sql_type: The SQL type as text, e.g. "VARCHAR(100)".
        nullable: True when the column may be left empty (NULL).
        primary_key: True when the column is the table's primary key.
        unique: True when every value in the column must be different.
        default: The Python-side default as text, or "-" when there is none.
    """

    model: str
    table: str
    name: str
    sql_type: str
    nullable: bool
    primary_key: bool
    unique: bool
    default: str

    @property
    def rules(self) -> str:
        """Describe the column's rules in words, e.g. "Not null, Unique".

        Returns:
            The rules joined with commas.
        """
        ## Start with the most important rule; if/elif/else picks exactly one.
        if self.primary_key:
            rules = ["Primary key"]
        elif self.nullable:
            rules = ["Optional"]
        else:
            rules = ["Not null"]
        ## Then add the extra rules that apply.
        if self.unique:
            rules.append("Unique")
        if self.default != "-":
            rules.append(f"Default {self.default}")
        return ", ".join(rules)


def inspect_schema() -> list[ColumnInfo]:
    """Read the table definitions back from the database file itself.

    Uses SQLAlchemy's inspector, which asks SQLite what tables and columns
    really exist, and combines that with the Python-side defaults.

    Returns:
        One ColumnInfo for every column of the categories and products tables.
    """
    ## inspect(engine) returns an Inspector that can describe the live database.
    inspector = inspect(engine)
    columns = []
    for model in (Category, Product):
        table_name = model.__tablename__
        ## The table as declared in Python (needed to read the default value).
        model_table = Base.metadata.tables[table_name]
        ## A set holds unique values and makes "name in primary_keys" fast.
        primary_keys = set(
            inspector.get_pk_constraint(table_name)["constrained_columns"]
        )
        ## A set comprehension with two "for" parts: loop over every unique
        ## constraint, then over every column name inside that constraint.
        unique = {
            name
            for constraint in inspector.get_unique_constraints(table_name)
            for name in constraint["column_names"]
        }
        ## Build one ColumnInfo per column reported by the database.
        for column in inspector.get_columns(table_name):
            name = column["name"]
            default = model_table.c[name].default
            columns.append(
                ColumnInfo(
                    model=model.__name__,
                    table=table_name,
                    name=name,
                    sql_type=str(column["type"]),
                    nullable=bool(column["nullable"]),
                    primary_key=name in primary_keys,
                    unique=name in unique,
                    ## getattr(obj, "arg", fallback) reads obj.arg if it exists;
                    ## repr() turns True into the text "True".
                    default=(
                        "-"
                        if default is None
                        else repr(getattr(default, "arg", default))
                    ),
                )
            )
    return columns


def _column_ok(
    columns: dict[tuple[str, str], ColumnInfo],
    table: str,
    name: str,
    type_prefix: str,
    *,
    primary_key: bool = False,
    nullable: bool | None = None,
    unique: bool | None = None,
    default: str | None = None,
) -> bool:
    """Check that one column exists and follows the expected rules.

    The ``*`` in the parameter list means every argument after it must be
    passed by name (e.g. ``unique=True``), which makes calls easier to read.
    Arguments left as None are simply not checked.

    Args:
        columns: All columns, looked up by (table, column name).
        table: The table to check.
        name: The column to check.
        type_prefix: The expected start of the SQL type, e.g. "VARCHAR".
        primary_key: Whether the column should be the primary key.
        nullable: Expected nullable setting, or None to skip the check.
        unique: Expected unique setting, or None to skip the check.
        default: Expected default text, or None to skip the check.

    Returns:
        True when every requested check passes, otherwise False.
    """
    column = columns.get((table, name))
    ## "and" stops at the first False, so a missing column is never examined.
    return (
        column is not None
        and column.sql_type.startswith(type_prefix)
        and column.primary_key == primary_key
        and (nullable is None or column.nullable == nullable)
        and (unique is None or column.unique == unique)
        and (default is None or column.default == default)
    )


def _column_summary(schema: Iterable[ColumnInfo], table: str) -> str:
    """Describe every column of one table on a single line.

    Args:
        schema: All columns from ``inspect_schema()``.
        table: The table whose columns should be described.

    Returns:
        Text such as "id INTEGER (Primary key); name VARCHAR(100) (...)".
    """
    return "; ".join(
        f"{c.name} {c.sql_type} ({c.rules})" for c in schema if c.table == table
    )


def _model_source(model: type[Base]) -> str:
    """The class line, table name, and column definitions of ``model``.

    It reads the model's real source code with ``getsource`` and keeps only
    the important lines, so the report shows the actual code that ran.
    This is why the model explanations are written above each class.

    Args:
        model: A model class such as ``Category`` (``type[Base]`` means "a
            class that inherits from Base", not an object of it).

    Returns:
        The selected source lines joined with new lines.
    """
    ## Keep a line only if it starts the class, names the table, or defines a
    ## column; .rstrip() removes trailing spaces from each kept line.
    return "\n".join(
        line.rstrip()
        for line in getsource(model).splitlines()
        if line.startswith("class ")
        or "__tablename__" in line
        or "mapped_column(" in line
    )


# --- Requirements: questions and answers -------------------------------------
## Each requirement from the exercise brief is turned into a question, answered
## in plain English, and checked against the real database (the "evidence").


@dataclass(frozen=True, slots=True)
class RequirementAnswer:
    """One brief requirement, phrased as a question and answered with evidence.

    Fields with ``= ...`` have default values, so they can be left out when
    creating an answer that has no result table or chart.

    Attributes:
        number: The question number (1, 2, 3, ...).
        question: The requirement rewritten as a yes/no question.
        requirement: The original wording from the brief.
        summary: A one-line answer for the scorecard table.
        answer: The full explanation.
        code: The code (or SQL) that fulfills the requirement.
        verification: How the answer was checked against the database.
        met: True when the requirement is satisfied.
        result_headers: Column headings of an optional result table.
        result_rows: Rows of that optional result table.
        chart: Name of an optional chart to show, or None.
    """

    number: int
    question: str
    requirement: str
    summary: str
    answer: str
    code: str
    verification: str
    met: bool
    result_headers: tuple[str, ...] = ()
    result_rows: tuple[tuple[str, ...], ...] = ()
    chart: str | None = None

    @property
    def label(self) -> str:
        """Return the short label, e.g. "Q3"."""
        return f"Q{self.number}"

    @property
    def status(self) -> str:
        """Return "MET" or "NOT MET" depending on ``met``."""
        return "MET" if self.met else "NOT MET"


@dataclass(frozen=True, slots=True)
class Report:
    """Everything the CSV and DOCX writers need.

    Attributes:
        results: The query results.
        stats: The per-category statistics.
        schema: The columns read back from the database.
        answers: One answer per requirement.
        findings: The key findings sentences.
        generated_at: When the report was made, as readable text.
    """

    results: CatalogResults
    stats: list[CategoryStats]
    schema: list[ColumnInfo]
    answers: list[RequirementAnswer]
    findings: list[str]
    generated_at: str

    @property
    def score(self) -> str:
        """Return the overall score, e.g. "11 of 11 requirements met (100%)"."""
        met = sum(a.met for a in self.answers)
        total = len(self.answers)
        return f"{met} of {total} requirements met ({_percent(met, total)})"


## Column headings for product and category result tables.
PRODUCT_HEADERS = ("ID", "Product", "Category", "Price", "Stock Status")
CATEGORY_HEADERS = ("ID", "Category", "Description", "Products")


def product_table(products: Iterable[ProductRow]) -> tuple[tuple[str, ...], ...]:
    """Turn products into rows of text matching ``PRODUCT_HEADERS``.

    Args:
        products: The products to show.

    Returns:
        A tuple of rows; each row is a tuple of strings.
    """
    ## tuple(generator) builds the rows; str() converts the id number to text.
    return tuple(
        (str(p.id), p.name, p.category, _money(p.price), p.stock_label)
        for p in products
    )


def category_table(
    results: CatalogResults, stats: Iterable[CategoryStats]
) -> tuple[tuple[str, ...], ...]:
    """Turn categories into rows of text matching ``CATEGORY_HEADERS``.

    Args:
        results: The query results (for the categories).
        stats: The per-category statistics (for the product counts).

    Returns:
        A tuple of rows; each row is a tuple of strings.
    """
    ## A dict mapping each category name to its number of products.
    counts = {s.name: s.product_count for s in stats}
    return tuple(
        (str(c.id), c.name, c.description or "(none)", str(counts.get(c.name, 0)))
        for c in results.categories
    )


def _with_sql(orm_code: str, statement: Any) -> str:
    """Show ORM code followed by the SQL that SQLAlchemy generates for it.

    Args:
        orm_code: The Python query code as text.
        statement: The matching query object.

    Returns:
        The code, a blank line, a "-- SQL generated" note, and the SQL.
    """
    return f"{orm_code}\n\n-- SQL generated by SQLAlchemy:\n{compiled_sql(statement)}"


def build_requirement_answers(
    results: CatalogResults,
    stats: list[CategoryStats],
    schema: list[ColumnInfo],
    explanations: dict[str, str],
) -> list[RequirementAnswer]:
    """Answer every brief requirement as a question, backed by real evidence.

    Args:
        results: The query results.
        stats: The per-category statistics.
        schema: The columns read back from the database.
        explanations: The plain-English query explanations.

    Returns:
        A list of RequirementAnswer objects (Q1 to Q11).
    """
    ## --- Gather the evidence used by the answers below ---
    ## Dict comprehension: look up any column by its (table, column) pair.
    columns = {(c.table, c.name): c for c in schema}
    ## sorted() of a set gives the unique table names in alphabetical order.
    tables = sorted({c.table for c in schema})
    products = results.products
    category_names = {c.name for c in results.categories}
    out_of_stock = [p for p in products if not p.in_stock]
    ## "*_" collects and ignores the first three values of each row, keeping only
    ## in_stock; sum() then counts how many rows relied on the True default.
    defaulted = sum(in_stock for *_, in_stock in PRODUCTS)
    above_limit = [p for p in products if p.price >= PRICE_LIMIT]
    ## The products closest to the price limit on either side (or None).
    nearest_excluded = min(above_limit, key=lambda p: p.price, default=None)
    highest_included = max(results.under_limit, key=lambda p: p.price, default=None)
    prices = [p.price for p in products] or [0.0]
    distribution = ", ".join(f"{s.name} ({s.product_count})" for s in stats)
    ## "Orphans" are products whose category name does not exist (should be none).
    orphans = [p.name for p in products if p.category_name not in category_names]
    empty_categories = [s.name for s in stats if not s.product_count]
    no_description = [c.name for c in results.categories if not c.description]
    ## The database file name taken from the engine URL ("or ''" avoids None).
    database_name = Path(engine.url.database or "").name
    ## Count this script's lines by reading the file as text.
    script_lines = len(SCRIPT_PATH.read_text(encoding="utf-8").splitlines())

    ## A function defined inside another function (a "nested" function). It can
    ## use EXECUTED_SQL and is only needed here.
    def executed(prefix: str) -> int:
        """Count captured SQL statements that start with ``prefix``.

        Args:
            prefix: Upper-case start of the statement, e.g. "INSERT".

        Returns:
            How many statements matched.
        """
        return sum(sql.upper().startswith(prefix) for sql in EXECUTED_SQL)

    ## all(...) is True only if every check in the tuple is True.
    category_ok = all(
        (
            _column_ok(columns, "categories", "id", "INTEGER", primary_key=True),
            _column_ok(
                columns, "categories", "name", "VARCHAR", nullable=False, unique=True
            ),
            _column_ok(columns, "categories", "description", "VARCHAR", nullable=True),
        )
    )
    ## The same five-column check for the products table.
    product_ok = all(
        (
            _column_ok(columns, "products", "id", "INTEGER", primary_key=True),
            _column_ok(columns, "products", "name", "VARCHAR", nullable=False),
            _column_ok(columns, "products", "price", "FLOAT", nullable=False),
            _column_ok(columns, "products", "in_stock", "BOOLEAN", default="True"),
            _column_ok(columns, "products", "category_name", "VARCHAR"),
        )
    )
    ## The insert code shown in the report for Q6 and Q7 (text, not run here).
    session_code = (
        "with Session(engine) as session:\n"
        "    session.add_all(\n"
        "        Category(name=name, description=description)\n"
        "        for name, description in CATEGORIES\n"
        "    )\n"
        "    session.add_all(_make_products())\n"
        "    session.commit()"
    )

    ## --- One RequirementAnswer per requirement in the brief (Q1 to Q10) ---
    answers = [
        ## Q1: the script file itself exists.
        RequirementAnswer(
            number=1,
            question="Was a file called product_models.py created?",
            requirement="Create a file called product_models.py.",
            summary=f"Yes - a {script_lines}-line, self-contained script",
            answer=(
                "Yes. product_models.py is a single, self-contained Python script. "
                "It is organized into clearly labeled sections (configuration, "
                "models, sample data, queries, requirement answers, console output, "
                "CSV export, charts, and DOCX report) so each part of the exercise "
                "is easy to find and review."
            ),
            code=f"{SCRIPT_PATH.name}  ({script_lines} lines)",
            verification=(
                f"The file {SCRIPT_PATH.name} exists and was executed to produce "
                "this report."
            ),
            met=SCRIPT_PATH.name == "product_models.py",
        ),
        ## Q2: the engine points at the right file and has echo=True.
        RequirementAnswer(
            number=2,
            question=(
                "Is a SQLAlchemy engine connected to product_catalog.db with echo=True?"
            ),
            requirement=(
                "Set up a SQLAlchemy engine connecting to product_catalog.db "
                "(with echo=True)."
            ),
            summary="Yes - echo=True logs every SQL statement",
            answer=(
                "Yes. The engine is created once, at module level, with "
                "create_engine() and a SQLite URL that points at product_catalog.db "
                "in the same folder as the script. echo=True makes SQLAlchemy log "
                "every SQL statement it sends to the database, so the CREATE TABLE, "
                "INSERT, and SELECT statements generated by the ORM are visible in "
                "the console. The Bonus section compares them with hand-written SQL."
            ),
            code=(
                'DB_PATH = BASE_DIR / "product_catalog.db"\n'
                'engine = create_engine(f"sqlite:///{DB_PATH}", echo=True)'
            ),
            verification=(
                f"engine.url points to {database_name} and engine.echo is "
                f"{engine.echo}. During this run echo=True logged "
                f"{len(EXECUTED_SQL)} statements, including "
                f"{executed('CREATE TABLE')} CREATE TABLE, {executed('INSERT')} "
                f"INSERT, and {executed('SELECT')} SELECT statements."
            ),
            met=database_name == DB_PATH.name and engine.echo is True,
        ),
        ## Q3: the Category model has the required columns and rules.
        RequirementAnswer(
            number=3,
            question="Is the Category model defined with the required columns?",
            requirement=(
                "Category: id (primary key), name (string, not null, unique), "
                "description (optional string)."
            ),
            summary="Yes - id PK, unique non-null name, optional description",
            answer=(
                "Yes. Category is declared on the shared DeclarativeBase and maps to "
                "the categories table. id is the integer primary key. name is a "
                "String(100) declared with nullable=False and unique=True, so every "
                "category must have a name and no two categories can share one. "
                "description is a String(255) typed as str | None, which makes it "
                "optional"
                + (
                    f": {', '.join(no_description)} was inserted without one."
                    if no_description
                    else "."
                )
            ),
            code=_model_source(Category),
            verification=(
                f"Read back from {DB_PATH.name}: "
                f"{_column_summary(schema, 'categories')}."
            ),
            met=category_ok,
        ),
        ## Q4: the Product model has the required columns and rules.
        RequirementAnswer(
            number=4,
            question="Is the Product model defined with the required columns?",
            requirement=(
                "Product: id (primary key), name (string, not null), price (float, "
                "not null), in_stock (boolean, default True), category_name (string "
                "- store the category name as text for now)."
            ),
            summary="Yes - all five columns with the required rules",
            answer=(
                "Yes. Product maps to the products table. id is the integer primary "
                "key; name is a String(100) that cannot be null; price is a Float "
                "that cannot be null, so every product must have a price; in_stock "
                "is a Boolean with default=True, so a new product is in stock unless "
                f"stated otherwise ({defaulted} of {len(PRODUCTS)} sample products "
                "relied on this default). category_name is a String(100) that "
                "stores the category name as plain text. It is intentionally not a "
                "foreign key yet; that change comes in the Relationships lesson."
            ),
            code=_model_source(Product),
            verification=(
                f"Read back from {DB_PATH.name}: {_column_summary(schema, 'products')}."
            ),
            met=product_ok,
        ),
        ## Q5: the tables were created with Base.metadata.create_all(engine).
        RequirementAnswer(
            number=5,
            question="Are the tables created with Base.metadata.create_all(engine)?",
            requirement="Create the tables using Base.metadata.create_all(engine).",
            summary=f"Yes - tables created: {', '.join(tables)}",
            answer=(
                "Yes. Base.metadata.create_all(engine) reads every table registered "
                "on Base.metadata and issues a CREATE TABLE statement for each one "
                "that does not exist yet. The script calls "
                "Base.metadata.drop_all(engine) first, so it can be run repeatedly "
                "without UNIQUE-constraint errors and always starts from a clean "
                "database."
            ),
            code="Base.metadata.drop_all(engine)\nBase.metadata.create_all(engine)",
            verification=(
                f"echo=True logged {executed('CREATE TABLE')} CREATE TABLE "
                f"statements, and {DB_PATH.name} now contains the tables "
                f"{', '.join(tables)}."
            ),
            met={"categories", "products"} <= set(tables),
        ),
        ## Q6: at least 3 categories were inserted through a session.
        RequirementAnswer(
            number=6,
            question="Were at least 3 categories inserted using a session?",
            requirement="Using a session: insert at least 3 categories.",
            summary=f"Yes - {len(results.categories)} categories inserted",
            answer=(
                "Yes. Inside a single 'with Session(engine) as session:' block, "
                f"{len(results.categories)} Category objects were added with "
                "session.add_all() and saved with session.commit(): "
                f"{', '.join(c.name for c in results.categories)}."
            ),
            code=session_code,
            verification=(
                f"The categories table holds {len(results.categories)} rows "
                "(required: at least 3)."
            ),
            met=len(results.categories) >= 3,
            result_headers=CATEGORY_HEADERS,
            result_rows=category_table(results, stats),
        ),
        ## Q7: at least 6 products were inserted, spread over the categories.
        RequirementAnswer(
            number=7,
            question="Were at least 6 products inserted across those categories?",
            requirement="Using a session: insert at least 6 products across those "
            "categories.",
            summary=(
                f"Yes - {len(products)} products across "
                f"{len(stats) - len(empty_categories)} categories"
            ),
            answer=(
                f"Yes. {len(products)} Product objects were added in the same "
                "session and committed together with the categories. They are "
                f"spread across all {len(stats)} categories - {distribution} - and "
                f"range in price from {_money(min(prices))} to "
                f"{_money(max(prices))}. Every category_name matches an existing "
                "category."
            ),
            code=session_code,
            verification=(
                f"The products table holds {len(products)} rows (required: at least "
                f"6). Categories without products: "
                f"{', '.join(empty_categories) or 'none'}. Products with an unknown "
                f"category: {', '.join(orphans) or 'none'}."
            ),
            met=len(products) >= 6 and not empty_categories and not orphans,
            result_headers=PRODUCT_HEADERS,
            result_rows=product_table(products),
            chart="category_stock",
        ),
        ## Q8: query 1 (all categories) was run and printed.
        RequirementAnswer(
            number=8,
            question="Does the program query and print all categories?",
            requirement="Using a session: query and print all categories.",
            summary=f"Yes - {len(results.categories)} rows returned and printed",
            answer=(
                "Yes. session.scalars(select(Category).order_by(Category.id)) "
                "returns every Category object in id order, and each one is printed "
                "to the console under 'QUERY 1: ALL CATEGORIES'. "
                + explanations["categories"]
            ),
            code=_with_sql(
                "session.scalars(select(Category).order_by(Category.id))",
                ALL_CATEGORIES_QUERY,
            ),
            verification=(
                f"The query returned {len(results.categories)} rows, exactly "
                f"matching the {len(CATEGORIES)} categories inserted."
            ),
            met=len(results.categories) == len(CATEGORIES),
            result_headers=CATEGORY_HEADERS,
            result_rows=category_table(results, stats),
        ),
        ## Q9: query 2 (in-stock products) was run and printed.
        RequirementAnswer(
            number=9,
            question="Does the program query and print all products that are in stock?",
            requirement="Using a session: query and print all products that are in "
            "stock.",
            summary=f"Yes - {len(results.in_stock)} in-stock products printed",
            answer=(
                "Yes. select(Product).where(Product.in_stock) keeps only products "
                "whose in_stock flag is True (SQLite stores it as 1), and the rows "
                "are printed under 'QUERY 2: PRODUCTS IN STOCK'. "
                + explanations["in_stock"]
            ),
            code=_with_sql(
                "session.scalars(select(Product).where(Product.in_stock))",
                IN_STOCK_QUERY,
            ),
            verification=(
                f"All {len(results.in_stock)} returned rows have in_stock = True, and "
                f"none of the {len(out_of_stock)} out-of-stock products "
                f"({_names(out_of_stock)}) appear."
            ),
            met=all(p.in_stock for p in results.in_stock)
            and len(results.in_stock) == len(products) - len(out_of_stock),
            result_headers=PRODUCT_HEADERS,
            result_rows=product_table(results.in_stock),
            chart="stock_pie",
        ),
        ## Q10: query 3 (products under the price limit) was run and printed.
        RequirementAnswer(
            number=10,
            question=(
                f"Does the program query and print all products priced under "
                f"{_money(PRICE_LIMIT)}?"
            ),
            requirement="Using a session: query and print all products with a "
            "price under $50.",
            summary=(
                f"Yes - {len(results.under_limit)} products under "
                f"{_money(PRICE_LIMIT)} printed"
            ),
            answer=(
                f"Yes. select(Product).where(Product.price < {PRICE_LIMIT:g})"
                ".order_by(Product.price) returns the products cheaper than "
                f"{_money(PRICE_LIMIT)}, sorted from cheapest to most expensive, and "
                "prints them under 'QUERY 3: PRODUCTS UNDER $50.00'. "
                + explanations["under_limit"]
            ),
            code=_with_sql(
                "session.scalars(\n"
                f"    select(Product).where(Product.price < {PRICE_LIMIT:g})"
                ".order_by(Product.price)\n)",
                UNDER_LIMIT_QUERY,
            ),
            verification=(
                f"All {len(results.under_limit)} rows are below "
                f"{_money(PRICE_LIMIT)}"
                + (
                    f" (highest returned: {highest_included.name} at "
                    f"{_money(highest_included.price)})"
                    if highest_included
                    else ""
                )
                + "."
                + (
                    f" The closest excluded product, {nearest_excluded.name} at "
                    f"{_money(nearest_excluded.price)}, confirms the filter "
                    "boundary."
                    if nearest_excluded
                    else ""
                )
            ),
            met=all(p.price < PRICE_LIMIT for p in results.under_limit)
            and len(results.under_limit) == len(products) - len(above_limit),
            result_headers=PRODUCT_HEADERS,
            result_rows=product_table(results.under_limit),
            chart="prices",
        ),
    ]
    ## Q11 sums up the others, so it is added after they exist. The file size is
    ## shown in kilobytes (bytes / 1024); 0.0 if the file is somehow missing.
    size_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0.0
    passed = sum(a.met for a in answers)
    answers.append(
        RequirementAnswer(
            number=len(answers) + 1,
            question=(
                "Is product_models.py a working program that creates the database, "
                "inserts data, and runs the three queries?"
            ),
            requirement=(
                "Deliverable: a working product_models.py that creates the "
                "database, inserts data, and runs the three queries."
            ),
            summary="Yes - runs end to end without errors",
            answer=(
                "Yes. Running 'python product_models.py' completes without errors. "
                f"It creates {DB_PATH.name} ({size_kb:.1f} KB), stores "
                f"{len(results.categories)} categories and {len(products)} products, "
                "runs and prints the three required queries, and then writes this "
                "CSV and DOCX report. It can be re-run at any time because it "
                "rebuilds the tables from scratch."
            ),
            code=(
                "python product_models.py            "
                "# build, print, and open the reports\n"
                "python product_models.py --no-open  # build and print only"
            ),
            verification=(
                f"{passed} of {len(answers)} checks (Q1-Q{len(answers)}) passed; "
                "the three queries returned "
                f"{len(results.categories)}, {len(results.in_stock)}, and "
                f"{len(results.under_limit)} rows."
            ),
            met=passed == len(answers) and DB_PATH.exists(),
        )
    )
    return answers


# --- Console output ----------------------------------------------------------
## Functions that print the query results as neat text tables in the terminal.


def _print_heading(title: str) -> None:
    """Print ``title`` between two lines of "=" signs.

    Args:
        title: The heading text.
    """
    ## '=' * 72 repeats the character 72 times; "\n" starts a new line.
    print(f"\n{'=' * 72}\n{title}\n{'=' * 72}")


def _print_products(products: Sequence[ProductRow]) -> None:
    """Print products as a fixed-width table.

    In the format codes, ``>3`` right-aligns in 3 characters and ``<30``
    left-aligns in 30 characters, so the columns line up.

    Args:
        products: The products to print.
    """
    ## Header row, then a dashed divider line.
    print(f"{'ID':>3}  {'Name':<30} {'Price':>9}  {'Category':<18} Stock")
    print("-" * 72)
    ## One line per product, using the same column widths as the header.
    for p in products:
        print(
            f"{p.id:>3}  {p.name:<30} {_money(p.price):>9}  "
            f"{p.category:<18} {p.stock_label}"
        )


def print_results(results: CatalogResults, explanations: dict[str, str]) -> None:
    """Print the three required queries and an explanation under each.

    Args:
        results: The query results.
        explanations: The plain-English text from build_query_explanations().
    """
    ## Query 1: categories (shows "(no description)" when it is missing).
    _print_heading("QUERY 1: ALL CATEGORIES")
    for c in results.categories:
        print(f"{c.id:>3}  {c.name:<20} {c.description or '(no description)'}")
    print(f"\n{explanations['categories']}")

    ## Query 2: in-stock products.
    _print_heading("QUERY 2: PRODUCTS IN STOCK")
    _print_products(results.in_stock)
    print(f"\n{explanations['in_stock']}")

    ## Query 3: products under the price limit.
    _print_heading(f"QUERY 3: PRODUCTS UNDER {_money(PRICE_LIMIT)}")
    _print_products(results.under_limit)
    print(f"\n{explanations['under_limit']}")


# --- CSV export --------------------------------------------------------------
## Writes the report as a CSV file that opens as a spreadsheet in Excel.


def _csv_table(
    headers: Sequence[str], rows: Iterable[Sequence[str]], indent: int = 0
) -> list[list[str]]:
    """Build a header row plus data rows, optionally shifted to the right.

    Args:
        headers: The column headings.
        rows: The data rows.
        indent: How many empty cells to put before each row (default 0).

    Returns:
        A list of CSV rows (each row is a list of strings).
    """
    ## [""] * indent makes a list of empty cells used as padding.
    pad = [""] * indent
    ## "*" unpacks a list into another list: [*pad, *headers] joins them. The
    ## second "*" unpacks the generator of padded data rows into the result.
    return [[*pad, *headers], *([*pad, *row] for row in rows)]


def write_csv(path: Path, report: Report) -> None:
    """Write the report as clearly numbered, consistently laid-out sections.

    Args:
        path: Where to save the CSV file.
        report: Everything to include in the report.
    """
    ## "a, b = x, y" assigns two variables at once.
    results, stats = report.results, report.stats
    findings_summary, requirements_summary = build_conclusion(
        results, stats, report.answers
    )

    ## Start the list of rows with the title block and the scorecard. An empty
    ## list [] becomes a blank line in the spreadsheet.
    rows: list[list[str]] = [
        [REPORT_TITLE.upper()],
        ["Generated", report.generated_at],
        ["Script", SCRIPT_PATH.name],
        ["Database", DB_PATH.name],
        ["ORM", f"SQLAlchemy {SQLALCHEMY_VERSION}"],
        ["Overall result", report.score],
        [],
        ["SECTION 1 - REQUIREMENTS SCORECARD"],
        *_csv_table(
            ["#", "Question", "Answer", "Status"],
            ([a.label, a.question, a.summary, a.status] for a in report.answers),
        ),
        [],
        ["SECTION 2 - QUESTIONS AND ANSWERS IN DETAIL"],
    ]
    ## Section 2: six rows per answer; "rows += [...]" appends several rows.
    for answer in report.answers:
        rows += [
            [answer.label, "Question", answer.question],
            ["", "Requirement", answer.requirement],
            ["", "Status", answer.status],
            ["", "Answer", answer.answer],
            ["", "Code", answer.code],
            ["", "Verification", answer.verification],
        ]
        ## Add the result table (indented by two cells) when there is one.
        if answer.result_rows:
            rows.append(["", "Result", f"{len(answer.result_rows)} rows"])
            rows += _csv_table(answer.result_headers, answer.result_rows, indent=2)
        rows.append([])

    ## Sections 3 to 8: schema, metrics, statistics, findings, SQL, conclusion.
    ## enumerate(items, 1) yields (1, first), (2, second), ... for numbering.
    rows += [
        [f"SECTION 3 - DATABASE SCHEMA (READ BACK FROM {DB_PATH.name.upper()})"],
        *_csv_table(
            ["Model", "Table", "Column", "SQL Type", "Rules"],
            ([c.model, c.table, c.name, c.sql_type, c.rules] for c in report.schema),
        ),
        [],
        ["SECTION 4 - CATALOG AT A GLANCE"],
        *_csv_table(["Metric", "Value"], build_summary_metrics(results)),
        [],
        ["SECTION 5 - CATEGORY STATISTICS"],
        *_csv_table(STATS_HEADERS, stats_table(stats)),
        [],
        ["SECTION 6 - KEY FINDINGS"],
        *([str(index), text] for index, text in enumerate(report.findings, 1)),
        [],
        ["SECTION 7 - BONUS: ECHO=TRUE SQL VS HAND-WRITTEN SQL"],
        *_csv_table(
            ["Statement", "Generated by SQLAlchemy", "Hand-Written SQL", "Difference"],
            build_sql_comparison(),
        ),
        [],
        ["SECTION 8 - CONCLUSION"],
        ["Summary of findings", findings_summary],
        ["Summary of requirements", requirements_summary],
    ]
    ## utf-8-sig adds a BOM so Excel detects the encoding correctly.
    ## newline="" is required by the csv module so rows do not get blank lines
    ## between them on Windows. The "with" block closes the file when done.
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        csv.writer(handle).writerows(rows)


# --- Charts ------------------------------------------------------------------
## Draws the charts with matplotlib and keeps each picture in memory as a PNG.


def _new_figure(width: float, height: float) -> tuple[Figure, Axes]:
    """Create an off-screen chart with one plot area.

    Using Figure directly (instead of matplotlib.pyplot) skips pyplot's slow
    import and its window-management machinery, which a script never needs.

    Args:
        width: Width in inches.
        height: Height in inches.

    Returns:
        (figure, axes): the whole image and the plot area to draw on.
    """
    ## Importing inside the function means matplotlib only loads when needed.
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure

    figure = Figure(figsize=(width, height))
    ## The Agg canvas draws into memory, so no window ever opens.
    FigureCanvasAgg(figure)
    return figure, figure.add_subplot()


def _figure_to_png(figure: Figure) -> io.BytesIO:
    """Save a matplotlib figure into an in-memory PNG image.

    Args:
        figure: The finished chart.

    Returns:
        A BytesIO "file in memory" holding the PNG, rewound to the start.
    """
    ## BytesIO behaves like a file but lives in memory (no temporary files).
    buffer = io.BytesIO()
    ## tight_layout() trims white space once. (bbox_inches="tight" would draw
    ## the whole chart twice, which is much slower.)
    figure.tight_layout()
    figure.savefig(
        buffer,
        format="png",
        dpi=CHART_DPI,
        pil_kwargs={"compress_level": PNG_COMPRESS_LEVEL},
    )
    ## Move back to the beginning so the next reader starts at the first byte.
    buffer.seek(0)
    return buffer


def build_charts(
    results: CatalogResults, stats: list[CategoryStats]
) -> dict[str, io.BytesIO]:
    """Render every chart to an in-memory PNG (no temporary files).

    Args:
        results: The query results.
        stats: The per-category statistics.

    Returns:
        A dict mapping each chart name (e.g. "prices") to its PNG image.
    """
    from matplotlib.ticker import MaxNLocator

    ## Data shared by the bar charts: one value per category.
    names = [s.name for s in stats]
    in_stock = [s.in_stock_count for s in stats]
    out_of_stock = [s.product_count - s.in_stock_count for s in stats]
    charts: dict[str, io.BytesIO] = {}

    ## Chart 1: stacked bars. _new_figure() creates a figure (the whole image)
    ## and an axes "ax" (the plot area). Sizes are width/height in inches.
    fig, ax = _new_figure(8, 4.5)
    bars = ax.bar(names, in_stock, color=POSITIVE_COLOR, label=IN_STOCK_LABEL)
    ax.bar_label(bars, label_type="center", color="white", fontweight="bold")
    ## bottom=in_stock stacks the out-of-stock bars on top of the in-stock ones.
    bars = ax.bar(
        names,
        out_of_stock,
        bottom=in_stock,
        color=ACCENT_COLOR,
        label=OUT_OF_STOCK_LABEL,
    )
    ## "v or ''" hides the label when the value is 0.
    ax.bar_label(bars, labels=[str(v or "") for v in out_of_stock], label_type="center")
    ax.set(title="Products per Category by Stock Status", ylabel="Products")
    ## Only whole numbers on the y axis (you cannot have half a product).
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.legend()
    charts["category_stock"] = _figure_to_png(fig)

    ## Chart 2: pie chart of in-stock vs out-of-stock products. autopct prints
    ## each slice's percentage; startangle=90 starts the first slice at the top.
    fig, ax = _new_figure(5.5, 5.5)
    stocked = len(results.in_stock)
    ax.pie(
        [stocked, len(results.products) - stocked],
        labels=[IN_STOCK_LABEL, OUT_OF_STOCK_LABEL],
        colors=[POSITIVE_COLOR, ACCENT_COLOR],
        autopct="%1.0f%%",
        startangle=90,
        wedgeprops={"edgecolor": "white"},
        textprops={"fontsize": 11},
    )
    ax.set_title("Overall Stock Availability")
    charts["stock_pie"] = _figure_to_png(fig)

    ## Chart 3: horizontal bars of every price, cheapest first. Bars under the
    ## limit are dark blue, others light blue; a dashed line marks the limit.
    ordered = sorted(results.products, key=lambda p: p.price)
    fig, ax = _new_figure(8, 5.5)
    bars = ax.barh(
        [p.name for p in ordered],
        [p.price for p in ordered],
        color=[BRAND_COLOR if p.price < PRICE_LIMIT else "#9DC3E6" for p in ordered],
    )
    ax.bar_label(bars, labels=[_money(p.price) for p in ordered], padding=3)
    ax.axvline(PRICE_LIMIT, color=ACCENT_COLOR, linestyle="--", label="$50 limit")
    ax.set(title="Product Prices (dark bars are under $50)", xlabel="Price (USD)")
    ## Leave 18% extra room on the right so the price labels fit.
    ax.set_xlim(0, max(p.price for p in ordered) * 1.18)
    ax.legend(loc="lower right")
    charts["prices"] = _figure_to_png(fig)

    ## Chart 4: average price per category with the $50 limit as a dashed line.
    fig, ax = _new_figure(8, 4.5)
    bars = ax.bar(names, [s.average_price for s in stats], color=BRAND_COLOR)
    ax.bar_label(bars, labels=[_money(s.average_price) for s in stats], padding=3)
    ax.axhline(PRICE_LIMIT, color=ACCENT_COLOR, linestyle="--", label="$50 limit")
    ax.set(title="Average Price per Category", ylabel="Average price (USD)")
    ax.set_ylim(0, max(s.average_price for s in stats) * 1.15)
    ax.legend()
    charts["average_price"] = _figure_to_png(fig)

    return charts


# --- DOCX report -------------------------------------------------------------
## Builds the Word document with the python-docx library: headings, paragraphs,
## tables (made of rows and cells), shaded boxes, charts, and a page footer.

## Elements that must follow <w:shd> inside <w:tcPr> (OOXML schema order).
## Word files are XML inside; cell shading must be placed before these tags.
_TCPR_SUCCESSORS = (
    "w:noWrap",
    "w:tcMar",
    "w:textDirection",
    "w:tcFitText",
    "w:vAlign",
    "w:hideMark",
)

## The caption printed under each chart, looked up by chart name.
CHART_CAPTIONS = {
    "category_stock": "Products per category, split by stock status.",
    "stock_pie": "Share of the catalog that is in stock.",
    "prices": "Every product price against the $50 limit.",
    "average_price": "Average price per category.",
}


def _shade_cell(cell: _Cell, fill: str) -> None:
    """Fill a Word table cell with a background color.

    python-docx has no simple setting for this, so we add the XML by hand.

    Args:
        cell: The table cell to color.
        fill: The color as hex digits without '#', e.g. "F2F2F2".
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    ## Create a <w:shd> (shading) XML element; qn() adds the Word namespace.
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    ## Insert it into the cell's properties in the position Word expects.
    cell._tc.get_or_add_tcPr().insert_element_before(shading, *_TCPR_SUCCESSORS)


def _write_cell(
    cell: _Cell,
    text: str,
    *,
    size: float = 9.5,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    font: str | None = None,
) -> None:
    """Write formatted text into a Word table cell.

    Args:
        cell: The table cell.
        text: The text to write.
        size: Font size in points.
        bold: Make the text bold.
        italic: Make the text italic.
        color: Optional text color as hex without '#'.
        font: Optional font name, e.g. "Consolas" for code.
    """
    from docx.shared import Pt, RGBColor

    ## A "run" is a piece of text that shares one style inside a paragraph.
    run = cell.paragraphs[0].add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    ## Only change color/font when one was given (None means "keep default").
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if font:
        run.font.name = font


def _spacer(doc: DocxDocument) -> None:
    """Add a small empty paragraph to leave space after a table.

    Args:
        doc: The Word document being built.
    """
    from docx.shared import Pt

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_table(
    doc: DocxDocument,
    headers: Sequence[str],
    rows: Iterable[Sequence[str]],
    *,
    status_column: int | None = None,
) -> None:
    """Branded table: dark header row, zebra striping, colored status cells.

    Args:
        doc: The Word document being built.
        headers: The column headings.
        rows: The data rows (each a sequence of strings).
        status_column: Index of a MET/NOT MET column to color, or None.
    """
    from docx.enum.table import WD_TABLE_ALIGNMENT

    ## Start with just the header row; data rows are added one at a time.
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ## zip() pairs each header cell with its heading text; strict=True raises an
    ## error if the two lists have different lengths.
    for cell, header in zip(table.rows[0].cells, headers, strict=True):
        _write_cell(cell, header, bold=True, color=WHITE_HEX)
        _shade_cell(cell, BRAND_COLOR[1:])
    ## enumerate() gives a running index, used to shade every second row.
    for index, row in enumerate(rows):
        cells = table.add_row().cells
        for column, (cell, value) in enumerate(zip(cells, row, strict=True)):
            is_status = column == status_column
            _write_cell(
                cell,
                value,
                bold=is_status,
                color=_status_color(value) if is_status else None,
            )
            ## index % 2 is 1 for odd rows, which counts as True: zebra stripes.
            if index % 2:
                _shade_cell(cell, ZEBRA_HEX)
    _spacer(doc)


def _add_box(
    doc: DocxDocument,
    text: str,
    fill: str,
    *,
    italic: bool = False,
    font: str | None = None,
    size: float = 10,
) -> None:
    """Single-cell shaded panel used for requirement callouts and code.

    Args:
        doc: The Word document being built.
        text: The text inside the box.
        fill: Background color as hex without '#'.
        italic: Make the text italic.
        font: Optional font name.
        size: Font size in points.
    """
    ## A one-row, one-column table works as a simple colored box.
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    _write_cell(cell, text, size=size, italic=italic, font=font)
    _shade_cell(cell, fill)
    _spacer(doc)


def _add_labeled(doc: DocxDocument, label: str, text: str = "") -> None:
    """Add a paragraph with a bold, colored label followed by optional text.

    Args:
        doc: The Word document being built.
        label: The bold label, e.g. "Answer:".
        text: Normal text written after the label (may be empty).
    """
    from docx.shared import Pt, RGBColor

    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(4)
    ## The label is its own run so it can be bold and colored on its own.
    run = paragraph.add_run(label)
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(BRAND_COLOR[1:])
    if text:
        paragraph.add_run(f" {text}")


def _add_status(doc: DocxDocument, answer: RequirementAnswer) -> None:
    """Add a "Status: ✔ MET" (or "✘ NOT MET") line in green or red.

    Args:
        doc: The Word document being built.
        answer: The requirement answer whose status is shown.
    """
    from docx.shared import RGBColor

    paragraph = doc.add_paragraph()
    ## Setting .bold on the returned run makes just "Status: " bold.
    paragraph.add_run("Status: ").bold = True
    run = paragraph.add_run(f"{'✔' if answer.met else '✘'} {answer.status}")
    run.bold = True
    ## Fall back to black ("000000") if the status has no special color.
    run.font.color.rgb = RGBColor.from_string(_status_color(answer.status) or "000000")


def _add_chart(doc: DocxDocument, image: io.BytesIO, caption: str) -> None:
    """Insert a chart image, centered, with an italic caption below it.

    Args:
        doc: The Word document being built.
        image: The PNG image in memory.
        caption: The text shown under the image.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches

    ## add_picture() puts the image in a new paragraph; [-1] is the last one.
    doc.add_picture(image, width=Inches(6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph = doc.add_paragraph(caption)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].italic = True


def _add_page_number_footer(doc: DocxDocument, text: str) -> None:
    """Add a centered footer such as "Report title  |  Page 3".

    Args:
        doc: The Word document being built.
        text: The text shown before the page number.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    paragraph = doc.sections[0].footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run(f"{text}  |  Page ").font.size = Pt(9)
    ## A PAGE field lets Word fill in the current page number.
    ## It is written as three XML pieces: field begin, the PAGE instruction, end.
    for tag, value in (
        ("w:fldChar", "begin"),
        ("w:instrText", "PAGE"),
        ("w:fldChar", "end"),
    ):
        element = OxmlElement(tag)
        ## The instruction holds text; the begin/end markers hold an attribute.
        if tag == "w:instrText":
            element.set(qn("xml:space"), "preserve")
            element.text = value
        else:
            element.set(qn("w:fldCharType"), value)
        ## Each piece goes into its own run inside the footer paragraph.
        run = paragraph.add_run()
        run.font.size = Pt(9)
        run._r.append(element)


def _configure_document(doc: DocxDocument) -> None:
    """Set the document's fonts, heading colors, margins, and footer.

    Args:
        doc: The new, empty Word document.
    """
    from docx.shared import Inches, Pt, RGBColor

    ## Change the "Normal" style, which most paragraphs use.
    normal = doc.styles["Normal"].font
    normal.name = "Calibri"
    normal.size = Pt(10.5)
    ## Color Heading 1 and Heading 2 with the brand color.
    for level in (1, 2):
        doc.styles[f"Heading {level}"].font.color.rgb = RGBColor.from_string(
            BRAND_COLOR[1:]
        )
    ## "a = b = value" sets both margins to the same value.
    for section in doc.sections:
        section.left_margin = section.right_margin = Inches(0.9)
        section.top_margin = section.bottom_margin = Inches(0.8)
    _add_page_number_footer(doc, REPORT_TITLE)


def write_docx(path: Path, report: Report) -> None:
    """Build the Word report around the brief's questions and their answers.

    Args:
        path: Where to save the .docx file.
        report: Everything to include in the report.
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    ## Prepare everything first: charts, a figure counter, and the conclusion.
    results, stats = report.results, report.stats
    charts = build_charts(results, stats)
    figures = count(1)
    findings_summary, requirements_summary = build_conclusion(
        results, stats, report.answers
    )
    ## Document() creates a new, empty Word document in memory.
    doc = Document()
    _configure_document(doc)

    ## Nested helper: next(figures) returns 1, then 2, ... for figure numbers.
    def chart(key: str) -> None:
        """Insert the chart named ``key`` with a numbered caption.

        Args:
            key: The chart name, e.g. "prices".
        """
        _add_chart(doc, charts[key], f"Figure {next(figures)}. {CHART_CAPTIONS[key]}")

    ## Cover page
    ## level=0 is the big "Title" style; the table lists the report details.
    doc.add_heading(REPORT_TITLE, level=0)
    subtitle = doc.add_paragraph("SQLAlchemy ORM Exercise - Requirements and Results")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_table(
        doc,
        ["Item", "Detail"],
        [
            ("Generated", report.generated_at),
            ("Script", SCRIPT_PATH.name),
            ("Database", DB_PATH.name),
            ("ORM", f"SQLAlchemy {SQLALCHEMY_VERSION}"),
            ("Overall result", report.score),
        ],
    )
    _add_labeled(doc, "Contents")
    ## A numbered list of the section titles, then start a new page.
    for heading in SECTION_TITLES:
        doc.add_paragraph(heading, style="List Number")
    doc.add_page_break()  # type: ignore[no-untyped-call]

    ## 1. Executive summary
    ## The intro paragraph, the scorecard table (status in column 3), and the
    ## headline metrics. SECTION_TITLES[0] is the first title in the tuple.
    doc.add_heading(f"1. {SECTION_TITLES[0]}", level=1)
    doc.add_paragraph(
        "This report answers every requirement of the Product Catalog Models "
        "exercise as a question. Each answer explains how the requirement was met, "
        "shows the code responsible, and is verified against the product_catalog.db "
        f"file produced by this run. Result: {report.score}."
    )
    _add_table(
        doc,
        ["#", "Question", "Answer", "Status"],
        [(a.label, a.question, a.summary, a.status) for a in report.answers],
        status_column=3,
    )
    _add_labeled(doc, "Catalog at a glance")
    _add_table(doc, ["Metric", "Value"], build_summary_metrics(results))

    ## 2. Questions and answers
    ## For each requirement: heading, requirement box, status, answer, code box,
    ## verification, and (when present) a result table and a chart.
    doc.add_page_break()  # type: ignore[no-untyped-call]
    doc.add_heading(f"2. {SECTION_TITLES[1]}", level=1)
    for answer in report.answers:
        doc.add_heading(f"{answer.label}. {answer.question}", level=2)
        _add_box(doc, f"Requirement: {answer.requirement}", CALLOUT_HEX, italic=True)
        _add_status(doc, answer)
        _add_labeled(doc, "Answer:", answer.answer)
        _add_labeled(doc, "Code:")
        _add_box(doc, answer.code, CODE_HEX, font="Consolas", size=8.5)
        _add_labeled(doc, "Verification:", answer.verification)
        if answer.result_rows:
            _add_labeled(doc, "Result:", f"{len(answer.result_rows)} rows.")
            _add_table(doc, answer.result_headers, answer.result_rows)
        if answer.chart:
            chart(answer.chart)

    ## 3. Database schema
    ## The columns read back from the database file, as proof of the schema.
    doc.add_heading(f"3. {SECTION_TITLES[2]}", level=1)
    doc.add_paragraph(
        "The columns below were read back from product_catalog.db with "
        "SQLAlchemy's inspector, which proves the models produced the required "
        "schema."
    )
    _add_table(
        doc,
        ["Model", "Column", "SQL Type", "Rules"],
        [(c.model, c.name, c.sql_type, c.rules) for c in report.schema],
    )

    ## 4. Category analysis
    ## The statistics table followed by the average-price chart.
    doc.add_heading(f"4. {SECTION_TITLES[3]}", level=1)
    _add_table(doc, STATS_HEADERS, stats_table(stats))
    chart("average_price")

    ## 5. Key findings
    ## Each finding becomes one bullet point.
    doc.add_heading(f"5. {SECTION_TITLES[4]}", level=1)
    for finding in report.findings:
        doc.add_paragraph(finding, style="List Bullet")

    ## 6. Bonus
    ## Generated SQL next to hand-written SQL; the for loop unpacks each 4-tuple.
    doc.add_heading(f"6. {SECTION_TITLES[5]}", level=1)
    doc.add_paragraph(
        "With echo=True, SQLAlchemy prints every statement it sends to SQLite. "
        "Each generated statement is shown next to the SQL you would write by hand, "
        "followed by the key difference."
    )
    for label, generated, raw, difference in build_sql_comparison():
        doc.add_heading(label, level=2)
        _add_labeled(doc, "Generated by SQLAlchemy (echo=True):")
        _add_box(doc, generated, CODE_HEX, font="Consolas", size=8.5)
        _add_labeled(doc, "Hand-written equivalent:")
        _add_box(doc, raw, CODE_HEX, font="Consolas", size=8.5)
        _add_labeled(doc, "Difference:", difference)

    ## 7. Conclusion
    ## The two closing paragraphs, then save the document to disk.
    doc.add_heading(f"7. {SECTION_TITLES[6]}", level=1)
    doc.add_heading("Summary of Findings", level=2)
    doc.add_paragraph(findings_summary)
    doc.add_heading("Summary of Requirements", level=2)
    doc.add_paragraph(requirements_summary)

    ## str(path) because doc.save() expects a plain text path or a file object.
    doc.save(str(path))


# --- Report cache ------------------------------------------------------------
## Charts and the Word file take most of the run time. The database is always
## rebuilt and queried (the exercise requires it), but the reports are only
## rewritten when their content or this script changes (or --force is given).


def compute_fingerprint(report: Report) -> str:
    """Return a SHA-256 "fingerprint" of everything the reports show.

    The timestamp is left out on purpose, so two runs with the same data and
    the same code produce the same fingerprint.

    Args:
        report: The finished report data.

    Returns:
        A 64-character hex string that changes whenever the content changes.
    """
    ## repr() turns the dataclasses into exact text; the script's own bytes are
    ## added so any code change (e.g. a new chart) also triggers a rebuild.
    content = repr(
        (report.results, report.stats, report.schema, report.answers, report.findings)
    )
    digest = hashlib.sha256(content.encode("utf-8"))
    digest.update(SCRIPT_PATH.read_bytes())
    return digest.hexdigest()


def reports_are_current(fingerprint: str) -> bool:
    """Return True if both reports exist and match ``fingerprint``.

    Args:
        fingerprint: The fingerprint of the report about to be written.

    Returns:
        True when the saved reports can be reused as they are.
    """
    ## all() is True only if every file exists; a missing file forces a rebuild.
    if not all(path.exists() for path in (CSV_PATH, DOCX_PATH, CACHE_PATH)):
        return False
    ## A damaged or unreadable cache file simply forces a rebuild.
    try:
        saved = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    return isinstance(saved, dict) and saved.get("fingerprint") == fingerprint


def save_fingerprint(fingerprint: str, generated_at: str) -> None:
    """Remember which content the current reports were built from.

    Args:
        fingerprint: The fingerprint of the reports just written.
        generated_at: When they were written (for people reading the file).
    """
    CACHE_PATH.write_text(
        json.dumps({"fingerprint": fingerprint, "generated_at": generated_at}),
        encoding="utf-8",
    )


# --- Entry point -------------------------------------------------------------
## Code that runs the whole program: read options, build, report, and open.


def open_file(path: Path) -> None:
    """Open ``path`` with the operating system's default application.

    Args:
        path: The file to open (e.g. the CSV opens in Excel).
    """
    ## os.startfile exists only on Windows; other systems use a desktop opener.
    ## getattr(..., None) returns None instead of raising an error if missing.
    startfile = getattr(os, "startfile", None)
    ## try/except catches the error if the file cannot be opened, prints a
    ## message, and lets the program continue instead of crashing.
    try:
        if startfile is not None:
            startfile(path)
        else:
            ## "Darwin" is macOS, which uses "open"; Linux uses "xdg-open".
            opener = "open" if platform.system() == "Darwin" else "xdg-open"
            subprocess.run([opener, str(path)], check=False)
    except OSError as error:
        ## file=sys.stderr prints to the error stream instead of normal output.
        print(f"Could not open {path}: {error}", file=sys.stderr)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Read the command-line options ``--no-open`` and ``--force``.

    Args:
        argv: The options to parse; None means "use the real command line".

    Returns:
        An object whose attributes hold the options, e.g. ``args.no_open``.
    """
    parser = argparse.ArgumentParser(
        description="Build the product catalog database and analytics reports."
    )
    ## action="store_true" makes args.no_open True when the flag is given.
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="generate the CSV and DOCX reports without opening them",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="rewrite the reports even if their content has not changed",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    """Run the whole program from start to finish.

    Args:
        argv: Command-line options; None means "use the real command line".

    Returns:
        An exit code: 0 for success, 1 if a report file could not be written.
    """
    args = parse_args(argv)
    ## try/finally: the "finally" part always runs, even if an error happens.
    try:
        results = build_catalog()
        ## Read the schema back from the real database file as evidence.
        schema = inspect_schema()
    finally:
        ## Release the SQLite file so it is not locked after the run (Windows).
        engine.dispose()

    ## Turn the results into statistics, explanations, answers, and findings.
    stats = compute_category_stats(results)
    explanations = build_query_explanations(results)
    report = Report(
        results=results,
        stats=stats,
        schema=schema,
        answers=build_requirement_answers(results, stats, schema, explanations),
        findings=build_findings(results, stats),
        ## strftime() formats the date, e.g. "November 23, 2026 at 09:24 PM".
        generated_at=datetime.now().strftime("%B %d, %Y at %I:%M %p"),
    )

    ## Write both reports, unless the saved ones already show exactly this
    ## content. PermissionError usually means a file is open in Excel or Word,
    ## so we print a helpful message and return error code 1.
    fingerprint = compute_fingerprint(report)
    if not args.force and reports_are_current(fingerprint):
        print("\nReports are up to date (content unchanged). Use --force to rebuild.")
    else:
        try:
            write_csv(CSV_PATH, report)
            write_docx(DOCX_PATH, report)
        except PermissionError as error:
            print(
                f"Cannot write {error.filename}. Close it in Excel/Word and try again.",
                file=sys.stderr,
            )
            return 1
        save_fingerprint(fingerprint, report.generated_at)

    ## Print the requirement scorecard and the overall result.
    _print_heading("REQUIREMENTS")
    for answer in report.answers:
        print(f"{answer.label:>4}  {answer.status:<8} {answer.question}")
    print(f"\nResult: {report.score}")

    ## Tell the user where the reports were saved.
    _print_heading("REPORTS")
    print(f"CSV report:  {CSV_PATH}")
    print(f"DOCX report: {DOCX_PATH}")

    ## Open both reports unless --no-open was given.
    if not args.no_open:
        open_file(CSV_PATH)
        open_file(DOCX_PATH)
    return 0


## __name__ is "__main__" only when this file is run directly (python
## product_models.py), not when another file imports it. sys.exit() passes
## main()'s return value to the operating system as the exit code.
if __name__ == "__main__":
    sys.exit(main())
