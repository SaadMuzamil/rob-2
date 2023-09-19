from robocorp.tasks import task
from robocorp import browser, http
import csv
from RPA.PDF import PDF
from time import sleep
from RPA.Archive import Archive


@task
def order_robot_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    open_webiste()
    login()
    downlaod_excel()
    go_to_order_page()
    close_model()
    import_excel_sheet()
    archive_receipts()
    # fill_and_submit_form()

def open_webiste():
    """Nevigate to the url"""
    browser.goto("https://robotsparebinindustries.com/")

def login():
    """Login to the website"""
    page = browser.page()
    page.fill("#username","maria")
    page.fill("#password","thoushallnotpass")
    page.click("button:text('Log in')")

def go_to_order_page():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_model():
    page = browser.page()
    page.click("button:text('OK')")

def fill_and_submit_form(row):
    """fill the form"""
    
    page = browser.page()
    page.select_option("#head",str(row["Head"]))
    page.set_checked("#id-body-"+str(row["Body"]),True)
    page.fill("input[type='number']",str(row["Legs"]))
    page.fill("#address",row["Address"])
    page.click("button:text('Order')")
    alert = page.query_selector(".alert-danger")
    if alert and alert.is_visible():
        print(alert)
        page1 = browser.page()
        page1.click("button:text('Order')")
        
    collect_result(row["Order number"])

    page.click("button:text('Order another robot')")

    close_model()
    


def downlaod_excel():
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)


def import_excel_sheet():
    with open('orders.csv') as csv_file:
        reader = csv.DictReader(csv_file,delimiter=",", quotechar='"')
        for row in reader:
            fill_and_submit_form(dict(row))


    


def collect_result(order_number):
    """Collect data"""

    page = browser.page()
    sales_results_html = page.locator("#order-completion").inner_html()
    pdf = PDF()
    pdf_file = pdf.html_to_pdf(sales_results_html, "output/receipts/sales_results"+order_number+".pdf")

    screen = page.locator("#robot-preview-image")
    screen_shoot = screen.screenshot(path="output/receipts/sales_summary"+order_number+".png")

    pdf.add_files_to_pdf(
        append=True,
        files=["output/receipts/sales_summary"+order_number+".png"],
        target_document="output/receipts/sales_results"+order_number+".pdf"
    )

   
def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('./output/receipts', 'receipts.zip')


