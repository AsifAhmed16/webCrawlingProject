import os
import xlwt
import requests

from bs4 import BeautifulSoup
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

QUANTITY = 6
MAIN_PAGE = 'https://shop.adidas.jp'
LIST_PAGE = MAIN_PAGE + '/item/?cat2Id=eoss22ss&order=10&gender=mens'


def get_product_details(product_tail, driver):
    # item details url
    item_url = MAIN_PAGE + product_tail

    driver.get(item_url)

    driver.execute_script("window.scrollTo(0, 2000);")
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "sizeDescription"))
        )
    except:
        pass
    driver.execute_script("window.scrollTo(0, 3000);")
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "BVRRRatingNormalOutOf"))
        )
    except:
        pass
    driver.execute_script("window.scrollTo(0, 4000);")
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "BVRRReviewDate"))
        )
    except:
        pass
    driver.execute_script("window.scrollTo(0, 6000);")

    soup = BeautifulSoup(driver.page_source, 'lxml')

    # breadcrumb - Category
    breadcrumb_ul = soup.find('ul', class_="breadcrumbList")
    breadcrumb_top = breadcrumb_ul.find('li', class_="top").text
    breadcrumb_li = breadcrumb_ul.find_all('li', class_="breadcrumbLink test-breadcrumbLink")
    breadcrumb_list = []
    for li in breadcrumb_li:
        breadcrumb_list.append(li.a.text)
    breadcrumb_tail = ' / '.join(breadcrumb_list)

    # Category name
    category = soup.find('span', class_="categoryName").text
    # Product name
    product = soup.find('h1', class_="itemTitle").text
    # Pricing
    pricing = soup.find('div', class_="articlePrice").p.span.text

    # Available Size
    available_size_div = soup.find('div', class_="test-sizeSelector css-10ei5xd")
    available_size = ""
    if available_size_div:
        available_size_li = available_size_div.find_all('li')
        size_list = []
        for li in available_size_li:
            size_list.append(li.text)
        available_size = ', '.join(size_list)

    """
    Sense of The Size
    Since it has no image/title to pick its value.
    """
    sense_of_the_size = "N/A"

    # Coordinated Products
    coordinated_products = []
    try:
        coordinated_div = soup.find('div', class_="coordinateItems")
        if coordinated_div:
            coordinated_li = coordinated_div.find_all('li', class_="css-1gzdh76")
            for li in coordinated_li:
                product_no = (li.find('div', class_="coordinate_image").img.attrs['src']).split("/")[3]
                coordinate_obj = {
                    'coordinated_product_name': li.find('div', class_="coordinate_image").img.attrs['alt'],
                    'pricing': li.find('div', class_="coordinate_price").text,
                    'product_number': product_no,
                    'image_url': (MAIN_PAGE + li.find('div', class_="coordinate_image").img.attrs['src']),
                    'product_page_url': (MAIN_PAGE + "/products/" + product_no + "/"),
                }
                coordinated_products.append(coordinate_obj)
    except:
        pass

    # Title of the description
    description_title = "N/A"
    try:
        description_title = soup.find('h4', class_="itemFeature heading test-commentItem-subheading").text
    except:
        pass
    # General Description of the Product
    general_description = "N/A"
    try:
        general_description = soup.find('div', class_="commentItem-mainText test-commentItem-mainText").text
    except:
        pass

    # General Description - Itemization
    itemization = []
    try:
        itemization_ul = soup.find('ul', class_="articleFeatures")
        itemization_li = itemization_ul.find_all('li')
        for li in itemization_li:
            itemization.append(li.text)
    except:
        pass

    # Size Chart
    size_chart_div = soup.find('div', class_="sizeDescription")
    size_chart = {}
    if size_chart_div:
        try:
            size_chart_heads = size_chart_div.find_all('table')[0]
            size_chart_data = size_chart_div.find_all('table')[1]
            table_head_tag = size_chart_heads.select('tr th')
            table_data_row_tag = size_chart_data.find_all('tr')
            table_head = []
            table_data = []
            for each in table_head_tag:
                text = each.get_text(separator=" ", strip=True)
                table_head.append(text)
            for each in table_data_row_tag:
                row_items = each.select('td')
                row_list = []
                for row in row_items:
                    row_list.append(row.text)
                row_data = ", ".join(row_list)
                table_data.append(row_data)
            size_chart = {table_head[i]: table_data[i] for i in range(len(table_head))}
        except:
            pass

    # Rating Scrapping
    rate_div = "N/A"
    try:
        rate_div = soup.find('div', class_='BVRRQuickTakeCustomWrapper')
    except:
        pass

    # Rating
    rating = "N/A"
    try:
        rating = rate_div.find('div', class_='BVRRRatingNormalOutOf').span.text
    except:
        pass

    # No of Review
    number_of_review = "N/A"
    try:
        review_div = rate_div.find('div', class_='BVRRBuyAgainContainer')
        number_of_review = review_div.find('span', class_='BVRRNumber BVRRBuyAgainTotal').text
    except:
        pass

    # Recommended Rate
    recommended_rate = "N/A"
    try:
        recommended_rate_div = rate_div.find('div', class_='BVRRRatingPercentage')
        recommended_rate = recommended_rate_div.find('span', class_='BVRRNumber').text
    except:
        pass

    # Review Rating
    fit = "N/A"
    length = "N/A"
    quality = "N/A"
    comfort = "N/A"
    try:
        review_rating_div = rate_div.find('div', class_='BVRRRatingContainerRadio')
        review_rating_odd_div = review_rating_div.find_all('div', class_='BVRRRatingEntry BVRROdd')
        review_rating_even_div = review_rating_div.find_all('div', class_='BVRRRatingEntry BVRREven')

        for div in review_rating_odd_div:
            fit_div = div.find('div', class_='BVRRRating BVRRRatingRadio BVRRRatingFit')
            quality_div = div.find('div', class_='BVRRRating BVRRRatingRadio BVRRRatingQuality')
            if fit_div:
                fit = fit_div.find('div', class_='BVRRRatingRadioImage').img.attrs['alt']
            elif quality_div:
                quality = quality_div.find('div', class_='BVRRRatingRadioImage').img.attrs['alt']
        for div in review_rating_even_div:
            length_div = div.find('div', class_='BVRRRating BVRRRatingRadio BVRRRatingLength')
            comfort_div = div.find('div', class_='BVRRRating BVRRRatingRadio BVRRRatingComfort')
            if length_div:
                length = div.find('div', class_='BVRRRatingRadioImage').img.attrs['alt']
            elif comfort_div:
                comfort = div.find('div', class_='BVRRRatingRadioImage').img.attrs['alt']
    except:
        pass

    review_rating = {
        'sense_of_fitting': fit,
        'appropriation_of_length': length,
        'quality_of_material': quality,
        'comfort': comfort,
    }

    # User Review Details
    user_review_details = []
    try:
        user_review_div = soup.find('div', class_='BVRRDisplayContentBody')
        user_review_div_list = user_review_div.findChildren("div", recursive=False)
        for div in user_review_div_list:
            div_by_id = div.find('div', class_='BVRRReviewDisplayStyle5Header')
            details = {
                'date': div_by_id.find('div', class_='BVRRReviewDateContainer').text,
                'rating': div_by_id.find('div', class_='BVRRRatingNormalImage').img.attrs['alt'],
                'review_title': div_by_id.find('div', class_='BVRRReviewTitleContainer').text
            }
            user_review_details.append(details)
    except:
        pass

    # KWs
    kws_list = []
    try:
        kws_div = soup.find('div', class_="itemTagsPosition")
        kws_a = kws_div.find_all('a')
        for kw in kws_a:
            kws_list.append(kw.text)
    except:
        pass

    # image urls
    image_list = []
    try:
        new_request = requests.get(item_url).text
        new_soup = BeautifulSoup(new_request, 'lxml')
        image_ul = new_soup.find('ul', class_="slider-list")
        image_li = image_ul.find_all('li')
        for li in image_li:
            image_list.append(MAIN_PAGE + li.button.div.img.attrs['src'])
    except:
        pass

    details = {
        'url': item_url,
        'breadcrumb': breadcrumb_top + " / " + breadcrumb_tail,
        'category': category,
        'image_urls': image_list,
        'product': product,
        'pricing': pricing,
        'available_size': available_size,
        'sense_of_the_size': sense_of_the_size,
        'coordinated_products': coordinated_products,
        'description_title': description_title,
        'general_description': general_description,
        'itemization': itemization,
        'size_chart': size_chart,
        'special_functions': "N/A",
        'rating': rating,
        'number_of_review': number_of_review,
        'recommended_rate': recommended_rate,
        'review_rating': review_rating,
        'user_review_details': user_review_details,
        'kws': ', '.join(kws_list),
    }

    return details


def get_product_list_contents():
    data = []
    page_no = 1

    file_dir = os.path.dirname(os.path.realpath(__file__))
    driver_path = os.path.join(file_dir, 'chromedriver')

    while page_no in range(1, QUANTITY):
        page_url = LIST_PAGE + f"&page={page_no}"
        html_text = requests.get(page_url).text
        soup = BeautifulSoup(html_text, 'lxml')
        html_cards = soup.find_all('div', class_="articleDisplayCard itemCardArea-cards test-card css-1lhtig4")
        for card in html_cards:
            try:
                driver = Chrome(executable_path=driver_path)
                card_obj = get_product_details(card.a.attrs['href'], driver)
                driver.quit()
                data.append(card_obj)
            except:
                pass
        page_no += 1

    return data


def generate_spreadsheet(all_data):
    workbook = xlwt.Workbook(encoding='utf-8')
    ws = workbook.add_sheet("items_list")

    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    header_style = xlwt.XFStyle()
    header_style.font = header_font
    font_style = xlwt.XFStyle()

    columns = ["S/L", "Product URL", "Breadcrumb", "Category", "Image Url's", "Product Name", "Pricing",
               "Available Size", "Sense of the Size", "Coordinated Product", "Description Title", "General Description",
               "Itemization", "Size Information", "Special Function", "Rating", "No of Reviews", "Recommended Rate",
               "Review Ratings", "User Review Details", "KWs"]

    row_num = 5
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], header_style)

    counter = 0
    for data in all_data:
        counter += 1
        row_num += 1
        ws.write(row_num, 0, counter)
        ws.write(row_num, 1, str(data['url']), font_style)
        ws.write(row_num, 2, str(data['breadcrumb']), font_style)
        ws.write(row_num, 3, str(data['category']), font_style)
        ws.write(row_num, 4, str(data['image_urls']), font_style)
        ws.write(row_num, 5, str(data['product']), font_style)
        ws.write(row_num, 6, str(data['pricing']), font_style)
        ws.write(row_num, 7, str(data['available_size']), font_style)
        ws.write(row_num, 8, str(data['sense_of_the_size']), font_style)
        ws.write(row_num, 9, str(data['coordinated_products']), font_style)
        ws.write(row_num, 10, str(data['description_title']), font_style)
        ws.write(row_num, 11, str(data['general_description']), font_style)
        ws.write(row_num, 12, str(data['itemization']), font_style)
        ws.write(row_num, 13, str(data['size_chart']), font_style)
        ws.write(row_num, 14, str(data['special_functions']), font_style)
        ws.write(row_num, 15, str(data['rating']), font_style)
        ws.write(row_num, 16, str(data['number_of_review']), font_style)
        ws.write(row_num, 17, str(data['recommended_rate']), font_style)
        ws.write(row_num, 18, str(data['review_rating']), font_style)
        ws.write(row_num, 19, str(data['user_review_details']), font_style)
        ws.write(row_num, 20, str(data['kws']), font_style)

    ws.write_merge(2, 3, 0, 5, "Product Details", xlwt.easyxf("font: bold on; align: vert centre, horiz center"))

    workbook.save("items.xls")


if __name__ == '__main__':
    all_data = get_product_list_contents()
    generate_spreadsheet(all_data)
