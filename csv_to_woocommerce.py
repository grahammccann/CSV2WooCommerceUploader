import csv
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from woocommerce import API
from dotenv import load_dotenv

load_dotenv()

SITE_URL = os.getenv("SITE_URL")
CONSUMER_KEY = os.getenv("CONSUMER_KEY")
CONSUMER_SECRET = os.getenv("CONSUMER_SECRET")


wcapi = API(
    url=SITE_URL,
    consumer_key=CONSUMER_KEY,
    consumer_secret=CONSUMER_SECRET,
    version="wc/v3",
    timeout=15
)

if not os.path.exists('uploaded.already'):
    with open('uploaded.already', 'w') as f:
        pass


def check_if_uploaded(code):
    with open('uploaded.already', 'r') as f:
        return code in f.read()


def mark_as_uploaded(code):
    with open('uploaded.already', 'a') as f:
        f.write(code + '\n')


def test_connection():
    response = wcapi.get("products")
    if response.status_code == 200:
        messagebox.showinfo("Success", "Connected successfully!")
    else:
        messagebox.showerror("Error",
                             f"Failed to connect. Status code: {response.status_code}. Message: {response.text}")


def generate_image_url(image_path):
    relative_path = image_path.replace("media/", "")
    return f"{SITE_URL}media/{relative_path}"


def markdown_to_html(description):
    description = description.replace("**", "<strong>", 1)
    description = description.replace("**", "</strong>", 1)

    lines = description.split("\n")
    for index, line in enumerate(lines):
        if line.startswith("* "):
            lines[index] = f"<li>{line[2:]}</li>"

    description = "\n".join(lines)
    if "<li>" in description:
        description = description.replace("<li>", "<ul>\n<li>", 1)
        description = description[::-1].replace(">il/<", ">lu/<", 1)[::-1]

    return description


def parse_category(category_str):
    """Parses the category string from the CSV into the desired format."""
    categories = category_str.strip().split(',')
    return ' / '.join(categories)


# Global variable to store WooCommerce categories
woocommerce_categories = {}


def initialize_woocommerce_categories():
    """Fetches all categories from WooCommerce and stores them in a global dictionary."""
    global woocommerce_categories
    response = wcapi.get("products/categories", params={"per_page": 100})  # Adjust per_page if you have more categories
    if response.status_code == 200:
        woocommerce_categories = {category['name']: category['id'] for category in response.json()}
    else:
        log_error(f"Failed to fetch categories. Error: {response.text}")


def get_or_create_category(category_name):
    global woocommerce_categories

    # Check if category already exists in our local cache
    if category_name in woocommerce_categories:
        return woocommerce_categories[category_name]

    # If not, try to create it
    response = wcapi.post("products/categories", {"name": category_name})
    if response.status_code == 201:
        category_id = response.json()["id"]
        woocommerce_categories[category_name] = category_id  # Update our local cache
        return category_id
    elif response.json().get("code") == "term_exists":  # If category already exists on WooCommerce
        # Fetch the existing category ID and update our local cache
        existing_category_id = response.json().get("data", {}).get("resource_id")
        if existing_category_id:
            woocommerce_categories[category_name] = existing_category_id
            return existing_category_id
    else:
        log_error(f"Failed to create category {category_name}. Error: {response.text}")
        return None


# Call this function once at the start of your script or when the GUI loads
initialize_woocommerce_categories()


def get_woocommerce_categories():
    """Fetches the list of categories from WooCommerce."""
    response = wcapi.get("products/categories")
    if response.status_code == 200:
        return {category['name']: category['id'] for category in response.json()}
    else:
        log_error(f"Failed to fetch categories. Error: {response.text}")
        return {}


def create_category(category_name):
    """Creates a new category in WooCommerce if it doesn't exist."""
    # Clean up the category name
    category_name = category_name.strip().rstrip('/').replace(", ", " / ")

    # Check if the category already exists
    existing_categories = get_woocommerce_categories()
    if category_name in existing_categories:
        return existing_categories[category_name]

    # If it doesn't exist, create it
    response = wcapi.post("products/categories", {"name": category_name})
    if response.status_code == 201:
        return int(response.json()['id'])  # Ensure the returned ID is an integer
    else:
        log_error(f"Failed to create category {category_name}. Error: {response.text}")
        return None


def search_products(query):
    # Clear the tree first
    for item in tree.get_children():
        tree.delete(item)

    # If the query is empty, reload all products
    if not query.strip():
        view_csv()
        return

    # Load the CSV and filter rows based on the search query
    csv_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not csv_path:
        return
    with open(csv_path, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        headers = next(reader)
        image_indices = [i for i, h in enumerate(headers) if h == "Image"]

        for row in reader:
            code = row[headers.index('Code')]
            first_image = row[image_indices[0]].strip()

            if check_if_uploaded(code) or not first_image:
                continue

            # Check if the query is in any of the columns of the row
            if any(query.lower() in cell.lower() for cell in row):
                images = [row[i] for i in image_indices]
                values = (
                    code, row[headers.index('Master Product Name')],
                    row[headers.index('Variant Name')],
                    row[headers.index('Product Description')], *images, row[headers.index('Manufacturer')],
                    row[headers.index('Category')], row[headers.index('Barcode')], row[headers.index('Colour')],
                    row[headers.index('Size 1')], row[headers.index('Size 2')], row[headers.index('RRP')],
                    row[headers.index('Stock')], row[headers.index('Brand + Product')],
                    row[headers.index('VAT Status')],
                    row[headers.index('Commodity Code')], row[headers.index('Country of Origin')]
                )
                tree.insert("", "end", values=values)


def upload_product_to_woocommerce(values, item):
    # Extracting data from the provided values
    code, master_product_name, variant_name, product_description, *image_paths, manufacturer, category, barcode, colour, size_1, size_2, rrp, stock, brand_product, vat_status, commodity_code, country_origin = values

    # Check if all image columns are empty
    image_paths = [img for img in image_paths if img.strip()]

    if not image_paths:
        return  # Skip rows with no images

    # Generate the image URL
    image_url = generate_image_url(image_paths[0])

    # Parse and get/create category
    parsed_category = parse_category(category)
    category_id = get_or_create_category(parsed_category)

    # Append size1, size2, and color to the product name if they are not empty
    product_name = master_product_name
    attributes_to_append = [size_1, size_2, colour]

    for attribute in attributes_to_append:
        if attribute and attribute.strip():
            product_name += f" ({attribute.strip()})"

    product_description = markdown_to_html(product_description)

    # Get or create the category ID
    category_id = get_or_create_category(category)
    if not category_id:
        log_error(f"Failed to get or create category {category} for product {code}.")
        return

    # Prepare product data for WooCommerce
    product_data = {
        "name": product_name,
        "type": "simple",
        "regular_price": rrp,
        "description": product_description,
        "short_description": product_name,
        "images": [{"src": image_url}],
        "sku": code,
        "categories": [{"id": category_id}]  # Use the category ID here
    }

    # Use the WooCommerce API to upload the product
    response = wcapi.post("products", product_data)

    if response.status_code != 201:
        error_message = f"Failed to upload product {code}. Error: {response.text}"
        log_error(error_message)
        # messagebox.showerror("Error", error_message)
    else:
        log_error(f"Successfully uploaded simple product {code}.")
        # messagebox.showinfo("Success", f"Successfully uploaded product {code}.")
        mark_as_uploaded(code)  # Mark the product as uploaded in the uploaded.already file
        tree.delete(item)  # Remove the row from the Treeview after successful upload


def log_error(message):
    with open('debug.log', 'a') as debug_file:
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        debug_file.write(f"[{timestamp}] {message}\n")


def upload_selected_rows_to_woocommerce():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showinfo("Info", "Please select rows to upload.")
        return

    for item in selected_items:
        values = tree.item(item, "values")
        upload_product_to_woocommerce(values, item)


def view_csv():
    csv_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not csv_path:
        return
    with open(csv_path, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        headers = next(reader)
        image_indices = [i for i, h in enumerate(headers) if h == "Image"]

        for row in reader:
            code = row[headers.index('Code')]
            first_image = row[image_indices[0]].strip()

            if check_if_uploaded(code) or not first_image:
                continue

            images = [row[i] for i in image_indices]

            values = (
                code, row[headers.index('Master Product Name')],
                row[headers.index('Variant Name')],
                row[headers.index('Product Description')], *images, row[headers.index('Manufacturer')],
                row[headers.index('Category')], row[headers.index('Barcode')], row[headers.index('Colour')],
                row[headers.index('Size 1')], row[headers.index('Size 2')], row[headers.index('RRP')],
                row[headers.index('Stock')], row[headers.index('Brand + Product')], row[headers.index('VAT Status')],
                row[headers.index('Commodity Code')], row[headers.index('Country of Origin')]
            )

            tree.insert("", "end", values=values)


root = tk.Tk()
root.title("WooCommerce Product Importer - graham23s@hotmail.com")
root.state('zoomed')

input_frame = tk.Frame(root, bg="#f0f0f0", padx=10, pady=10)
input_frame.pack(pady=20, fill=tk.X)

btn_style = {
    'font': ('Arial', 12, 'bold'),
    'padx': 20,
    'pady': 10
}

btn_view = tk.Button(input_frame, text="View CSV", command=view_csv, bg="#4CAF50", fg="white", **btn_style)
btn_view.grid(row=0, column=0, padx=20, pady=10)

btn_upload_selected = tk.Button(input_frame, text="Upload Selected Rows to WooCommerce",
                                command=upload_selected_rows_to_woocommerce, bg="#FF5722", fg="white", **btn_style)
btn_upload_selected.grid(row=0, column=1, padx=20, pady=10)

btn_test_connection = tk.Button(input_frame, text="Test Connection", command=test_connection, bg="#FFC107", fg="white",
                                **btn_style)
btn_test_connection.grid(row=0, column=2, padx=20, pady=10)

# Add this after your existing input_frame widgets
search_label = tk.Label(input_frame, text="Search:", font=('Arial', 12))
search_label.grid(row=0, column=3, padx=20, pady=10)

search_entry = tk.Entry(input_frame, font=('Arial', 12))
search_entry.grid(row=0, column=4, padx=20, pady=10)

btn_search = tk.Button(input_frame, text="Search", bg="#3F51B5", fg="white", font=('Arial', 12, 'bold'),
                       command=lambda: search_products(search_entry.get()))
btn_search.grid(row=0, column=5, padx=20, pady=10)

input_frame.grid_columnconfigure(0, weight=1)
input_frame.grid_columnconfigure(1, weight=1)
input_frame.grid_columnconfigure(2, weight=1)

columns = [
    "Code", "Master Product Name", "Variant Name", "Product Description",
    "Image_1", "Image_2", "Image_3", "Image_4", "Image_5", "Image_6", "Image_7", "Image_8", "Image_9", "Image_10",
    "Manufacturer", "Category", "Barcode", "Colour", "Size 1",
    "Size 2", "RRP", "Stock", "Brand + Product", "VAT Status", "Commodity Code",
    "Country of Origin"
]

frame = tk.Frame(root)
frame.pack(pady=20, fill=tk.BOTH, expand=True)

tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode='extended')
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, stretch=tk.YES)

vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
vsb.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=vsb.set)

hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
hsb.pack(side=tk.BOTTOM, fill=tk.X)
tree.configure(xscrollcommand=hsb.set)

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

root.mainloop()
