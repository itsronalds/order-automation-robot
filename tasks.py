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
    open_RobotSpareBin_website()
    download_excel_file()
    fill_the_form()
    archive_receipts()


def open_RobotSpareBin_website():
    """Open RobotSpareBin Website"""
    browser.goto('https://robotsparebinindustries.com/#/robot-order')


def download_excel_file():
    """Download orders excel file"""
    http = HTTP()
    http.download(
        url='https://robotsparebinindustries.com/orders.csv', overwrite=True)


def get_orders():
    """Get orders from the excel file"""
    csv = Tables()
    orders = csv.read_table_from_csv(path='orders.csv')

    results = []

    for order in orders:
        results.append(order)

    return results


def close_annoying_modal():
    """Close website modal"""
    page = browser.page()
    page.click('button:text("OK")')


def fill_the_form():
    """Complete form orders"""
    orders = get_orders()

    page = browser.page()

    for order in orders:
        close_annoying_modal()

        # complete form
        page.select_option('#head', order['Head'])
        page.check(f'#id-body-{order["Body"]}')
        page.fill(
            'input[placeholder="Enter the part number for the legs"]', order['Legs'])
        page.fill('#address', order['Address'])
        page.click('#preview')
        page.click('#order')

        while page.query_selector('.alert.alert-danger') is not None:
            page.click('#order')

        pdf_file = store_receipt_as_pdf(order['Order number'])
        screenshot = screenshot_robot(order['Order number'])
        embed_screenshot_to_receipt(screenshot, pdf_file)

        page.click('#order-another')


def store_receipt_as_pdf(order_number):
    """Export receipt data as a PDF"""
    page = browser.page()
    receipt_result_html = page.locator('#receipt').inner_html()

    pdf_url = f'output/{order_number}.pdf'

    pdf = PDF()
    pdf.html_to_pdf(receipt_result_html, pdf_url)

    return pdf_url


def screenshot_robot(order_number):
    """Take a robot preview screenshot"""
    page = browser.page()

    image_url = f'output/{order_number}.png'

    page.locator(
        '#robot-preview-image').screenshot(path=image_url)

    return image_url


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Insert the screenshot in the PDF file"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot, source_path=pdf_file, output_path=pdf_file)


def archive_receipts():
    """Create a ZIP file with all receipts"""
    archive = Archive()
    archive.archive_folder_with_zip(
        folder='./output', archive_name='receipts.zip', include='*.pdf')
