import time
from time import sleep
import logging
import xlwt
import re
from selenium import webdriver
from selenium.webdriver.common.by import By


# the reason to sleep is referred to bilibili
# the reason to retry is referred to 51cto
# the standardized call is:
# retry(lambda: ..., logger = logger)
def retry(func, max_try = 10, interval = 0.5, logger = None):
    for index_try in range(max_try):
        sleep(interval)
        try:
            return func()
        except:
            if logger:
                logger.info("Encounter an error(%d/%d), try again." % (index_try + 1, max_try))
    return func()


def xls_write(content, xls_file):
    workbook = xlwt.Workbook(encoding= 'utf-8')
    worksheet = workbook.add_sheet("sheet 1")

    if isinstance(content, dict):
        for row_index, row_key in enumerate(content):
            row_value = content[row_key]
            col_index = 0
            if isinstance(row_key, tuple):
                for cell_key in row_key:
                    worksheet.write(row_index, col_index, cell_key)
                    col_index += 1
            else:
                worksheet.write(row_index, col_index, row_key)
                col_index += 1
            if isinstance(row_value, dict):
                for cell_value in row_value.values():
                    worksheet.write(row_index, col_index, cell_value)
                    col_index += 1
            else:
                worksheet.write(row_index, col_index, row_value)
    elif isinstance(content, set):
        for row_index, row_content in enumerate(content):
            if isinstance(row_content, tuple):
                for col_index, cell_content in enumerate(row_content):
                    worksheet.write(row_index, col_index, cell_content)
            else:
                worksheet.write(row_index, 0, row_content)
    else:
        raise TypeError

    workbook.save(xls_file + ".xls")


if __name__ == "__main__":
    logging.basicConfig(level = logging.INFO)
    log_file = 'log_%s.log' % time.strftime('%Y%m%d%H%M')
    log_handler = logging.FileHandler(log_file, mode = 'w')
    log_formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d]: %(message)s")
    log_handler.setFormatter(log_formatter)
    logger = logging.getLogger(__name__)
    logger.addHandler(log_handler)

    movie_dict = dict()
    type_set = set()
    movie_to_imdb = dict()
    movie_to_type = set()
    director_dict = dict()
    movie_to_director = set()
    actor_dict = dict()
    character_dict = dict()
    company_dict = dict()

    # download Edge webdriver from the following website:
    # https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    wb = webdriver.Edge()
    wb.get("https://movie.douban.com/top250")

    for page_index in range(10):
        for film_index in range(25):
            # wb.find_elements_by_css_selector("") in version 3.141.0
            # wb.find_elements(By.CSS_SELECTOR, "") in latest version
            # with reference to: https://blog.csdn.net/BigDataPlayer/article/details/125549949:
            div_list = wb.find_elements(By.CSS_SELECTOR, "div.info")
            div = div_list[film_index]

            movie_name = div.text.split()[0]
            logger.info("Start to crawl movie %d on page %d: %s." % (film_index + 1, page_index + 1, movie_name))

            if movie_name in movie_to_imdb.keys():
                logger.info("This movie has been recorded.")
                continue

            div_text = div.find_element(By.CLASS_NAME, "bd").text.split("\n")[1]
            movie_year = int(div_text.split()[0])
            movie_type = div_text.split("/")[2].strip().split()

            film_link = div.find_element(By.CSS_SELECTOR, "span.title")
            retry(lambda: film_link.click(), logger = logger)

            div_text = wb.find_element(By.CSS_SELECTOR, "div#info").text
            movie_imdb = re.search(r"IMDb: (\w+)", div_text).group(1)
            movie_length = int(re.search(r"片长: (\d+)", div_text).group(1))

            movie_summary = "\n".join(wb.find_element(By.CSS_SELECTOR, "div.related-info").text.split("\n")[1:])

            movie_dict[movie_imdb] = {"电影名称": movie_name, "电影发行年份": movie_year,
                                      "电影长度": movie_length, "电影情节概要": movie_summary}
            movie_to_imdb[movie_name] = movie_imdb

            for type_name in movie_type:
                type_set.add(type_name)
                movie_to_type.add((movie_imdb, type_name))

            for director_link in wb.find_element(By.CSS_SELECTOR, "span.attrs").find_elements(By.CSS_SELECTOR, "a"):
                director_name = director_link.text
                logger.info("Start to crawl director: %s." % director_name)

                retry(lambda: director_link.click(), logger = logger)
                div_text = wb.find_element(By.CSS_SELECTOR, "div.info").text
                director_imdb = re.search(r"imdb编号: (\w+)", div_text).group(1)
                director_birthday = div_text.split("\n")[2].split()[1]

                director_dict[director_imdb] = {"导演姓名": director_name, "导演出生日期": director_birthday}
                movie_to_director.add((movie_imdb, director_imdb))
                wb.back()

            for actor_index in range(len(wb.find_elements(By.CSS_SELECTOR, "li.celebrity"))):
                actor_link = wb.find_elements(By.CSS_SELECTOR, "li.celebrity")[actor_index]
                # it looks awkward here
                # but otherwise it pops an error

                div_text = actor_link.text
                if div_text.split()[1] == "导演":
                    continue

                actor_name = div_text.split()[0]
                character_name = " ".join(div_text.split("\n")[1].split()[1:])
                logger.info("Start to crawl actor: %s." % actor_name)

                retry(lambda: actor_link.click(), logger = logger)
                div_text = wb.find_element(By.CSS_SELECTOR, "div.info").text
                actor_imdb = re.search(r"imdb编号: (\w+)", div_text).group(1)
                actor_birthday = div_text.split("\n")[2].split()[1]

                actor_dict[actor_imdb] = {"演员姓名": actor_name, "演员出生日期": actor_birthday}
                character_dict[(movie_imdb, actor_imdb)] = character_name
                wb.back()

            wb.back()

        retry(lambda: wb.find_element(By.CSS_SELECTOR, "div.paginator").find_elements(By.CSS_SELECTOR, "a")[-1].click(),
              logger = logger)
        logger.info("Now in next page: %d." % page_index + 2)

    xls_write(movie_dict, "movie")
    xls_write(type_set, "type")
    xls_write(movie_to_type, "movie_to_type")
    xls_write(director_dict, "director")
    xls_write(movie_to_director, "movie_to_director")
    xls_write(actor_dict, "actor")
    xls_write(character_dict, "character")
    xls_write(company_dict, "company")