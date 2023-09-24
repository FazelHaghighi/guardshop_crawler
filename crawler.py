import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

# Create a Chrome WebDriver instance
driver = webdriver.Chrome()

# URL of the website
url = "https://gardshop.ir/product/caseprint/"

# Initialize an empty list to store data
data = []

try:
    # Open the website
    driver.get(url)

    # Locate the select elements for Brand, Model, and Material
    brand_select = Select(driver.find_element(By.NAME, "attribute_%d8%a8%d8%b1%d9%86%d8%af-%da%af%d9%88%d8%b4%db%8c"))
    model_select = Select(driver.find_element(By.NAME, "attribute_%d9%85%d8%af%d9%84-%da%af%d9%88%d8%b4%db%8c"))
    material_select = Select(driver.find_element(By.NAME, "attribute_%d8%ac%d9%86%d8%b3"))

    # Get all brand options
    brand_options = [option.get_attribute("value") for option in brand_select.options if option.get_attribute("value")]

    # Get all model options
    model_options = [option.get_attribute("value") for option in model_select.options if option.get_attribute("value")]

    # Get all material options
    material_options = [option.get_attribute("value") for option in material_select.options if option.get_attribute("value")]

    # Loop through brand options
    for brand_name in brand_options:
        brand_select.select_by_value(brand_name)
        time.sleep(2)  # Add a delay to allow the page to load

        # Loop through model options
        for model_name in model_options:
            try:
                model_select.select_by_value(model_name)
                time.sleep(2)  # Add a delay to allow the page to load

                # Loop through material options
                for material_name in material_options:
                    try:
                        # Select the material if it's available
                        material_select.select_by_value(material_name)
                        time.sleep(2)  # Add a delay to allow the page to load

                        # Find the stock element
                        try:
                            stock_element = driver.find_element(By.CSS_SELECTOR, '.stock.in-stock')
                            stock = stock_element.text
                        except NoSuchElementException:
                            stock = "0"

                        print(f"Brand: {brand_name}, Model: {model_name}, Material: {material_name}, Stock: {stock}")

                        # Append data to the list
                        data.append([brand_name, model_name, material_name, stock])

                    except NoSuchElementException:
                        print(f"Stock not found for Brand: {brand_name}, Model: {model_name}, Material: {material_name}")
                        continue

                    except Exception as e:
                        print(f"An error occurred: {str(e)}")
                        break  # Break out of the material loop on any error

                # Go back to the default material selection
                material_select.select_by_value("")  # Select the default option to go back
                time.sleep(2)  # Add a delay to allow the page to load

            except NoSuchElementException:
                print(f"Model not found for Brand: {brand_name}, Model: {model_name}")
                continue

            except Exception as e:
                print(f"An error occurred: {str(e)}")
                break  # Break out of the model loop on any error

        # Go back to the default model selection
        model_select.select_by_value("")  # Select the default option to go back
        time.sleep(2)  # Add a delay to allow the page to load

    # Go back to the default brand selection
    brand_select.select_by_value("")  # Select the default option to go back
    time.sleep(2)  # Add a delay to allow the page to load

finally:
    # Create a pandas DataFrame from the collected data
    df = pd.DataFrame(data, columns=["Brand", "Model", "Material", "Stock"])

    # Export the DataFrame to an Excel file
    df.to_excel("phone_guards_stock.xlsx", index=False)

    # Quit the webdriver
    driver.quit()
