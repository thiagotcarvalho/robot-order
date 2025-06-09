from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)

    download_csv_file()
    orders = get_orders()
    
    if orders:
        navigate_to_robot_order_website()
    
        for order in orders:
            process_order(order)
            
        archive_receipts()
    
def navigate_to_robot_order_website():
    """
    Open the robot order website.
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
    if browser.page().locator("button:has-text('OK')").is_visible():
        print("A modal is visible, closing it...")
        close_annoying_modal()
    
def close_annoying_modal():
    """
    Close the annoying modal.
    """
    page = browser.page()
    page.click("button:has-text('OK')")
    
def download_csv_file():
    """
    Download the CSV file.
    """
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv",
                  overwrite=True)

def get_orders():
    """
    Get the orders.
    """
    data = Tables()
    orders = data.read_table_from_csv(
        "orders.csv",
        columns=["Order number", "Head", "Body", "Legs", "Address"])
    return orders

def fill_and_submit_the_form(order) -> bool:
    """
    Fill and submit the form.
    """
    print("Filling and submitting the form for order "
          f"{order['Order number']}...")
    page = browser.page()
    
    page.select_option(selector="select#head", value=str(order["Head"]))
    page.click(f"input#id-body-{order['Body']}")
    page.fill(selector="input[placeholder='Enter the part number for the legs']", value=str(order["Legs"]))
    page.fill(selector="input#address", value=order["Address"])
    page.click(selector="button:has-text('Preview')")    
    page.click(selector="button:has-text('ORDER')")
    
    for _ in range(3):
        if page.locator("div#receipt").is_visible():
            create_pdf(order["Order number"])
            print("Order successful, clicking ORDER ANOTHER ROBOT button...")
            page.click(selector="button:has-text('ORDER ANOTHER ROBOT')")
            return True
        else:
            print("Order failed, clicking ORDER button again...")
            page.click(selector="button:has-text('ORDER')")
    return False

def process_order(order) -> None:
    """
    Process the order.
    """
    if fill_and_submit_the_form(order):
        close_annoying_modal()
    else:
        print(f"Order {order['Order number']} failed, retrying...")

def store_receipt_as_pdf(order_number) -> str:
    """
    Store the receipt as a PDF file.
    """
    page = browser.page()
    receipt_html = page.locator("div#receipt").inner_html()
    
    pdf = PDF()
    pdf.html_to_pdf(receipt_html,
                    f"output/receipts/RobotSpareBin_Order_{order_number}.pdf")
    
    return f"output/receipts/RobotSpareBin_Order_{order_number}.pdf"

def screenshot_robot(order_number):
    """
    Screenshot the robot.
    """
    page = browser.page()
    page.screenshot(path=f"output/screenshots/RobotSpareBin_Order_{order_number}.png")
    return f"output/screenshots/RobotSpareBin_Order_{order_number}.png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """
    Embed the screenshot to the receipt.
    """
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot],
                         target_document=pdf_file,
                         append=True)
    
def create_pdf(order_number):
    """
    Create the receipt PDF with the screenshot and the receipt.
    """
    print(f"Creating PDF for order {order_number}...")
    
    print(f"Storing receipt as PDF for order {order_number}...")
    receipt_pdf_path = store_receipt_as_pdf(order_number)
    
    print(f"Screenshotting robot for order {order_number}...")
    screenshot_path = screenshot_robot(order_number)
    
    print(f"Embedding screenshot to receipt for order {order_number}...")
    embed_screenshot_to_receipt(screenshot_path,
                                receipt_pdf_path)
    
def archive_receipts():
    """
    Archive the receipts.
    """
    archive = Archive()
    archive.archive_folder_with_zip(
        folder="output/receipts",
        archive_name="output/receipts.zip")