import logging
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("log.txt", mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def get_config():
    """Reads hierarchical configuration from Config.txt."""
    config = {'products': {}}
    with open('Config.txt', 'r', encoding='utf-8') as f:
        page_context = None
        current_product = None
        for line in f:
            line = line.rstrip()
            if not line.strip():
                continue

            if line.startswith('Page '):
                page_context = line.strip()
                current_product = None
                continue

            if page_context == 'Page 2: Configuration':
                indent_level = len(line) - len(line.lstrip(' '))
                parts = line.strip().split()
                
                if len(parts) >= 2:
                    quantity_str = parts[-1]
                    product_name = " ".join(parts[:-1])
                    
                    try:
                        quantity = int(''.join(filter(str.isdigit, quantity_str)))
                        
                        if indent_level == 0:  # Main product
                            config['products'][product_name] = {
                                'quantity': quantity,
                                'accessories': {}
                            }
                            current_product = product_name
                        elif current_product and indent_level > 0:  # Accessory
                            if current_product in config['products']:
                                config['products'][current_product]['accessories'][product_name] = quantity
                                
                    except (ValueError, IndexError):
                        pass  # Ignore lines that don't parse correctly
            elif ':' in line:
                key, value = line.split(':', 1)
                config[key.strip()] = value.strip()
    return config

def get_credentials():
    """Reads username and password from Account.txt."""
    username = None
    password = None
    with open('Account.txt', 'r', encoding='utf-8') as f:
        for line in f:
            if '：' in line:
                key, value = line.split('：', 1)
                if key.strip() == 'username':
                    username = value.strip()
                elif key.strip() == 'password':
                    password = value.strip()
    return username, password

def click_tab_with_retry(driver, wait, tab_id, tab_name):
    """Waits for overlay to disappear and clicks a tab with retries."""
    for attempt in range(3):
        try:
            # Wait for any loading overlay to disappear
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.blockUI.blockOverlay")))
            
            # Find and click the tab
            tab_element = wait.until(EC.element_to_be_clickable((By.ID, tab_id)))
            tab_element.click()
            
            logging.info(f"Successfully clicked '{tab_name}' tab on attempt {attempt + 1}.")
            return True
        except Exception as e:
            logging.warning(f"Attempt {attempt + 1} to click '{tab_name}' tab failed: {e}")
            time.sleep(1) # Wait a bit before retrying
    
    logging.error(f"Failed to click '{tab_name}' tab after multiple attempts.")
    return False


def main():
    """Main function to perform login."""
    # Suppress webdriver_manager logs
    logging.getLogger('webdriver_manager').setLevel(logging.WARNING)



    username, password = get_credentials()
    if not username or not password:
        logging.error("Username or password not found in Account.txt")
        return

    # Setup webdriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    wait = WebDriverWait(driver, 20) # Increased wait time

    try:
        logging.info("Opening login page...")
        driver.get("https://iconfig-cloud.h3c.com/iconfig/Index")

        logging.info("Selecting language...")
        english_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".englishDiv a")))
        english_link.click()
        
        wait.until(EC.text_to_be_present_in_element((By.ID, "lblLoginTitle"), "Welcome Channel user to login the H3C Configurator"))
        logging.info("Language selected: English")

        logging.info("Clicking 'H3C user' tab...")
        h3c_user_tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#InternalUser a")))
        h3c_user_tab.click()
        logging.info("'H3C user' tab clicked.")

        wait.until(EC.visibility_of_element_located((By.ID, "userAccounts")))
        
        logging.info("Entering credentials...")
        username_field = wait.until(EC.element_to_be_clickable((By.ID, "userAccounts")))
        password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
        
        username_field.send_keys(username)
        password_field.send_keys(password)
        logging.info("Credentials entered.")

        logging.info("Clicking login button...")
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "login_submit")))
        login_button.click()

        logging.info("Login successful!")
        
        driver.maximize_window()
        logging.info("Browser window maximized.")

        logging.info("Waiting 10 seconds for page to load after login...")
        time.sleep(10)
        
        # Click the Quotation menu using JavaScript to ensure the click is registered
        logging.info("Attempting to click Quotation menu...")
        quotation_menu = wait.until(EC.element_to_be_clickable((By.ID, "myH3CQuotation")))
        driver.execute_script("arguments[0].click();", quotation_menu)
        logging.info("Clicked the Quotation menu.")
        time.sleep(10)

        # Wait for the 'New' button to be clickable on the new page
        logging.info("Waiting for 'New' button...")
        new_button = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "New")))
        new_button.click()
        logging.info("Clicked 'New' button.")

        # Get config and fill form
        config = get_config()
        
        quotation_name_input = wait.until(EC.visibility_of_element_located((By.ID, "quoterName")))
        quotation_name_input.send_keys(config.get("Quotation name"))
        logging.info(f"Entered Quotation name: {config.get('Quotation name')}")

        logging.info("Clicking country dropdown...")
        country_dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-id='countryCode']")))
        country_dropdown_button.click()

        country_name = config.get("Country")
        logging.info(f"Selecting country: {country_name}")
        country_option_xpath = f"//div[contains(@class, 'dropdown-menu')]//span[text()='{country_name}']"
        country_option = wait.until(EC.element_to_be_clickable((By.XPATH, country_option_xpath)))
        country_option.click()
        logging.info(f"Selected Country: {country_name}")

        if config.get("Is U.S. ECCN needed") == "Yes":
            eccn_radio = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='eccn'][value='1']")))
            eccn_radio.click()
            logging.info("Selected 'Is U.S. ECCN needed': Yes")
        else:
            eccn_radio = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='eccn'][value='0']")))
            eccn_radio.click()
            logging.info("Selected 'Is U.S. ECCN needed': No")

        logging.info("Basic information filled.")

        logging.info("Clicking 'Save' button...")
        save_button = wait.until(EC.element_to_be_clickable((By.ID, "quotation_sava_btn")))
        save_button.click()
        logging.info("Clicked 'Save' button.")

        time.sleep(2)

        logging.info("Clicking 'Configuration' tab...")
        configuration_tab = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "config_page")))
        configuration_tab.click()
        logging.info("Clicked 'Configuration' tab.")

        time.sleep(10) # Wait for configuration page to load

        # Add products from config
        for product_name, product_data in config.get('products', {}).items():
            quantity = product_data['quantity']
            accessories = product_data.get('accessories', {})
            is_parts_type = product_name.startswith('WA')
            is_standard_type = product_name.startswith(('S', 'F'))

            if is_standard_type:
                logging.info(f"Product {product_name} is a 'Standard' type. Clicking 'Standard' tab.")
                if not click_tab_with_retry(driver, wait, "normCfg", "Standard"):
                    continue
            elif is_parts_type:
                logging.info(f"Product {product_name} is a 'Parts' type. Clicking 'Parts' tab.")
                if not click_tab_with_retry(driver, wait, "partsCfg", "Parts"):
                    continue
            else:
                logging.warning(f"Product {product_name} has an unknown type. Assuming 'Standard' and proceeding.")
                if not click_tab_with_retry(driver, wait, "normCfg", "Standard"):
                    continue

            time.sleep(2) # Wait for tab content to load

            logging.info(f"Adding product {product_name}...")

            # 1. Search for the product
            logging.info(f"Searching for product: {product_name}")
            search_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder*='for multiple conditi']")))
            search_input.clear()
            search_input.send_keys(product_name)
            
            search_button = wait.until(EC.element_to_be_clickable((By.ID, "searchBtn")))
            search_button.click()
            logging.info("Search button clicked.")
            
            time.sleep(3) # Wait for search results
            
            # 2. Find the product row and get the Product Code (if standard)
            product_row_xpath = f"//tr[contains(., '{product_name}')]"
            product_row = wait.until(EC.visibility_of_element_located((By.XPATH, product_row_xpath)))

            product_code = None
            if not is_parts_type:
                product_code_element = product_row.find_element(By.XPATH, ".//td[2]")
                product_code = product_code_element.text
                logging.info(f"Captured Product Code for Standard type: {product_code}")

            # Before adding, get the current set of config names if it's a Parts type
            if is_parts_type:
                existing_configs = {el.get_attribute('configname') for el in driver.find_elements(By.XPATH, "//input[@name='checkList']")}
                logging.info(f"Existing config names before adding Part: {existing_configs}")

            # Select the checkbox for the product
            logging.info(f"Selecting checkbox for {product_name}")
            product_checkbox = product_row.find_element(By.XPATH, ".//input[@type='checkbox']")
            product_checkbox.click()
            logging.info("Product checkbox selected.")
            
            # 3. Click the 'Add' button (↓Add)
            logging.info("Clicking 'Add' button.")
            add_button = wait.until(EC.element_to_be_clickable((By.ID, "addToTable2")))
            add_button.click()
            logging.info("'Add' button clicked.")
            
            # 4. Click the 'OK' button
            logging.info("Clicking 'OK' button.")
            ok_button = wait.until(EC.element_to_be_clickable((By.ID, "ok_button")))
            ok_button.click()
            logging.info("'OK' button clicked.")

            time.sleep(3) # Wait for the main list to update

            # Find the row and click 'Edit'
            try:
                edit_button_xpath = ''
                if is_parts_type:
                    logging.info("Finding new config name for Parts type...")
                    # Find the new config name by comparing with the old set
                    new_configs = {el.get_attribute('configname') for el in driver.find_elements(By.XPATH, "//input[@name='checkList']")}
                    newly_added_configs = new_configs - existing_configs
                    if newly_added_configs:
                        new_config_name = newly_added_configs.pop()
                        logging.info(f"Found newly added Part config name: {new_config_name}")
                        edit_button_xpath = f"//tr[.//input[@name='checkList' and @configname='{new_config_name}']]//a[@class='editConfig pointer']"
                    else:
                        logging.error("Could not identify the newly added Part's config name.")
                        continue # Skip to next product
                else: # Standard type
                    logging.info(f"Finding row for standard product code {product_code} to edit.")
                    edit_button_xpath = f"//tr[.//input[@name='checkList' and contains(@configname, '{product_code}')]]//a[@class='editConfig pointer']"

                edit_button = wait.until(EC.element_to_be_clickable((By.XPATH, edit_button_xpath)))
                
                driver.execute_script("arguments[0].scrollIntoView(true);", edit_button)
                time.sleep(1)
                edit_button.click()
                logging.info(f"Clicked 'Edit' for the last added product.")

                time.sleep(2) # Wait for the edit dialog to appear

                # --- Configure group editing dialog ---
                # 1. Edit Config name
                config_name_input = wait.until(EC.visibility_of_element_located((By.ID, "configName")))
                config_name_input.clear()
                config_name_input.send_keys(product_name)
                logging.info(f"Set 'Config name' to '{product_name}'.")

                # 2. Edit Sets
                sets_input = wait.until(EC.visibility_of_element_located((By.NAME, "siteNum")))
                sets_input.clear()
                sets_input.send_keys(str(quantity))
                logging.info(f"Set 'Sets' to '{quantity}'.")

                # 3. Click OK to save changes
                dialog_ok_button = wait.until(EC.element_to_be_clickable((By.ID, "ok_button")))
                dialog_ok_button.click()
                logging.info("Clicked 'OK' in the group editing dialog.")

                # If accessories exist, enter detail page to add them
                if accessories:
                    logging.info(f"Entering detail page for {product_name} to add accessories.")
                    # From the screenshot, the link is an <a> tag with class 'showConfig'.
                    detail_link_xpath = f"//a[@class='showConfig' and contains(., '{product_name}')]"
                    detail_link = wait.until(EC.element_to_be_clickable((By.XPATH, detail_link_xpath)))
                    # Use JavaScript click to avoid potential interception
                    driver.execute_script("arguments[0].click();", detail_link)
                    logging.info(f"Clicked product name to enter detail page.")
                    time.sleep(5)  # Wait for detail page to load

                    # On the Components page, find and click the link ending with "#1" to go to the accessory selection page.
                    logging.info("On Components page, looking for link ending with '#1'.")
                    try:
                        switch_link_id = f"node_title__{product_code}_0"
                        switch_link = wait.until(EC.element_to_be_clickable((By.ID, switch_link_id)))
                        driver.execute_script("arguments[0].click();", switch_link)
                        logging.info("Clicked the switch link ending with '#1'.")
                        logging.info("Waiting for accessory page to load...")
                        wait.until(EC.presence_of_element_located((By.ID, "allzhankai")))
                        logging.info("Accessory page loaded.")

                        # Click 'Expand all'
                        try:
                            logging.info("Waiting for 'Expand all' button to be clickable.")
                            expand_all_button = WebDriverWait(driver, 20).until(
                                EC.element_to_be_clickable((By.ID, "expand_all"))
                            )
                            expand_all_button.click()
                            logging.info("Clicked 'Expand all'.")
                            time.sleep(2) # Wait for sections to expand
                        except Exception as e:
                            logging.error(f"Could not find or click 'Expand all': {str(e)}")
                            driver.save_screenshot("expand_all_error.png")

                        # Process each accessory
                        for acc_name, acc_qty in accessories.items():
                            cleaned_acc_name = acc_name.strip().lstrip('-').strip()
                            try:
                                logging.info(f"Processing accessory: '{cleaned_acc_name}'. Will select max value from dropdown or use configured quantity '{acc_qty}'.")

                                # 1. Find and click the accessory row to make it editable
                                # 使用更精确的XPath定位器
                                accessory_row_xpath = f"//tr[contains(@class, 'item_tr') and .//td[.//span[normalize-space(text())='{cleaned_acc_name}']]]"
                                logging.info(f"Waiting for accessory row: {cleaned_acc_name}")
                                
                                # 增加等待时间，确保元素完全加载
                                wait = WebDriverWait(driver, 20)
                                accessory_row = wait.until(EC.presence_of_element_located((By.XPATH, accessory_row_xpath)))
                                item_tr_id = accessory_row.get_attribute('id')
                                logging.info(f"Found accessory row for {cleaned_acc_name} with id: {item_tr_id}")
                                
                                # 确保元素在视图中并可点击
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", accessory_row)
                                time.sleep(2)  # 等待滚动完成
                                
                                # 1. 点击数量单元格以触发下拉列表
                                try:
                                    # 尝试多种可能的数量单元格定位器
                                    quantity_cell_xpath_list = [
                                        f"//tr[@id='{item_tr_id}']//td[contains(@class, 'item_qty')]"
                                    ]
                                    
                                    quantity_cell = None
                                    for qty_xpath in quantity_cell_xpath_list:
                                        try:
                                            logging.info(f"Trying quantity cell xpath: {qty_xpath}")
                                            quantity_cell = wait.until(EC.element_to_be_clickable((By.XPATH, qty_xpath)))
                                            logging.info(f"Found quantity cell using xpath: {qty_xpath}")
                                            break
                                        except TimeoutException:
                                            continue
                                    
                                    if not quantity_cell:
                                        raise Exception("Could not find quantity cell with any of the attempted xpaths")
                                    
                                    # 直接点击数量单元格，优先使用ActionChains
                                    try:
                                        logging.info(f"Attempting to click quantity cell for '{cleaned_acc_name}' with ActionChains.")
                                        ActionChains(driver).move_to_element(quantity_cell).click().perform()
                                        logging.info("ActionChains click on quantity cell successful.")
                                    except Exception as e:
                                        logging.warning(f"ActionChains click on quantity cell failed: {e}. Falling back to JavaScript click.")
                                        driver.execute_script("arguments[0].click();", quantity_cell)
                                        logging.info("Clicked quantity cell with JavaScript.")
                                    
                                    time.sleep(1) # 短暂等待，以防下拉菜单有动画
                                    
                                    # 等待行进入编辑状态
                                    try:
                                        wait.until(lambda driver: 'editing' in driver.find_element(By.XPATH, f"//tr[@id='{item_tr_id}']").get_attribute('class'))
                                        logging.info(f"Row {item_tr_id} is now in editing state")
                                    except TimeoutException:
                                        logging.warning(f"Row {item_tr_id} did not enter editing state, continuing anyway")

                                    # 2. 根据item_tr_id判断使用弹出菜单还是直接输入
                                    try:
                                        # 判断item_tr_id是否包含'License'，以决定处理方式
                                        if 'License' in item_tr_id:
                                            # 直接输入方式
                                            logging.info(f"'{item_tr_id}' contains 'License', using direct input method.")
                                            
                                            # 终极方法V2：通过获取活动元素来定位输入框
                                            logging.info(f"'{item_tr_id}' contains 'License', using active element strategy.")
                                            time.sleep(1) # 等待JS创建input并聚焦

                                            # 1. 直接获取当前页面的活动元素
                                            quantity_input = driver.switch_to.active_element
                                            if quantity_input and quantity_input.tag_name == 'input':
                                                logging.info(f"Successfully got active element: {quantity_input.get_attribute('outerHTML')}")
                                                # 2. 使用JS设值并触发事件
                                                js_script = """
                                                var input = arguments[0];
                                                var value = arguments[1];
                                                input.value = value;
                                                var event_input = new Event('input', { bubbles: true });
                                                var event_change = new Event('change', { bubbles: true });
                                                input.dispatchEvent(event_input);
                                                input.dispatchEvent(event_change);
                                                """
                                                driver.execute_script(js_script, quantity_input, str(acc_qty))
                                                logging.info(f"Set quantity to '{acc_qty}' and dispatched events for License item.")
                                                time.sleep(0.5)

                                                # 3. 发送Enter键确认
                                                from selenium.webdriver.common.keys import Keys
                                                quantity_input.send_keys(Keys.ENTER)
                                                logging.info("Sent Enter key to finalize input for License item.")
                                            else:
                                                raise Exception("Failed to get active element or it was not an input field.")
                                        else:
                                            # 尝试弹出菜单方式，如果失败则尝试直接输入
                                            logging.info(f"'{item_tr_id}' does not contain 'License', trying popup menu first.")
                                            popup_menu_success = False
                                            try:
                                                # 等待弹出菜单出现
                                                popup_menu_xpath = f"//div[@id='action_div' and contains(@class, 'popup-menu')]//a[text()='{acc_qty}']"
                                                target_option = WebDriverWait(driver, 3).until(
                                                    EC.element_to_be_clickable((By.XPATH, popup_menu_xpath))
                                                )
                                                
                                                # 点击目标选项
                                                target_option.click()
                                                logging.info(f"Successfully clicked popup menu option '{acc_qty}'")
                                                popup_menu_success = True
                                                
                                                # 发送Enter键确认
                                                from selenium.webdriver.common.keys import Keys
                                                target_option.send_keys(Keys.ENTER)
                                                logging.info("Sent Enter key to confirm popup menu selection")
                                                
                                            except Exception as e:
                                                logging.info(f"Popup menu approach failed: {e}. Trying direct input instead.")
                                            
                                            # 如果弹出菜单方式失败，尝试直接输入
                                            if not popup_menu_success:
                                                logging.info(f"Switching to direct input for quantity '{acc_qty}'...")
                                                
                                                # 终极方法V2：通过获取活动元素来定位输入框
                                                logging.info(f"Switching to active element strategy for quantity '{acc_qty}'...")
                                                time.sleep(1) # 等待JS创建input并聚焦

                                                # 1. 直接获取当前页面的活动元素
                                                quantity_input = driver.switch_to.active_element
                                                if quantity_input and quantity_input.tag_name == 'input':
                                                    logging.info(f"Successfully got active element: {quantity_input.get_attribute('outerHTML')}")
                                                    # 2. 使用JS设值并触发事件
                                                    js_script = """
                                                    var input = arguments[0];
                                                    var value = arguments[1];
                                                    input.value = value;
                                                    var event_input = new Event('input', { bubbles: true });
                                                    var event_change = new Event('change', { bubbles: true });
                                                    input.dispatchEvent(event_input);
                                                    input.dispatchEvent(event_change);
                                                    """
                                                    driver.execute_script(js_script, quantity_input, str(acc_qty))
                                                    logging.info(f"Set quantity to '{acc_qty}' and dispatched events.")
                                                    time.sleep(0.5)

                                                    # 3. 发送Enter键确认
                                                    from selenium.webdriver.common.keys import Keys
                                                    quantity_input.send_keys(Keys.ENTER)
                                                    logging.info("Sent Enter key to finalize input.")
                                                else:
                                                    raise Exception("Failed to get active element or it was not an input field.")
                                        
                                        # 等待行变为selected状态以确认操作成功
                                        WebDriverWait(driver, 5).until(
                                            lambda d: 'selected' in d.find_element(By.XPATH, f"//tr[@id='{item_tr_id}']").get_attribute('class')
                                        )
                                        logging.info(f"Row {item_tr_id} is now in selected state, confirming quantity update")

                                    except Exception as e:
                                        logging.error(f"Failed to select quantity for '{cleaned_acc_name}': {e}")
                                        driver.save_screenshot(f'error_screenshot_select_qty_{cleaned_acc_name}_{int(time.time())}.png')
                                        continue # 继续处理下一个附件

                                except Exception as e:
                                    logging.error(f"Failed to perform quantity selection for '{cleaned_acc_name}': {e}")
                                    driver.save_screenshot(f'error_screenshot_perform_qty_selection_{cleaned_acc_name}_{int(time.time())}.png')
                                    continue # 继续处理下一个附件

                                # 4. Wait for the row to become 'selected' to confirm the action
                                selected_row_xpath = f"{accessory_row_xpath}[contains(@class, 'selected')]"
                                logging.info(f"Waiting for row to become selected: {selected_row_xpath}")
                                wait.until(EC.presence_of_element_located((By.XPATH, selected_row_xpath)))
                                logging.info(f"Successfully selected quantity for '{cleaned_acc_name}'.")
                                time.sleep(1) # Pause before next accessory

                            except Exception as e:
                                logging.error(f"Could not process accessory '{cleaned_acc_name}': {str(e)}")
                                driver.save_screenshot(f"error_acc_{cleaned_acc_name.replace(' ', '_')}.png")
                                # Continue to the next accessory instead of crashing

                        # After configuring accessories, save and go back to the main configuration page.
                        logging.info("Finished configuring accessories. Returning to main configuration page.")
                        
                        # 保存配置
                        logging.info("Step 3: Saving configuration...")
                        try:
                            save_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, "h3c_save_config"))
                            )
                            save_button.click()
                            logging.info("Successfully clicked save button")
                            time.sleep(2)  # 等待保存完成
                        except Exception as e:
                            logging.error(f"Error clicking save button: {e}")
                        
                        # 返回主列表
                        logging.info("Step 4: Returning to main list...")
                        try:
                            back_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, "back_list"))
                            )
                            back_button.click()
                            logging.info("Successfully clicked back button, returned to main list")
                            time.sleep(2)  # 等待页面加载
                        except Exception as e:
                            logging.error(f"Error clicking back button: {e}")
                            raise

                    except Exception as e:
                        logging.error(f"Could not process accessories on Components page for {product_name}: {e}")
                        driver.save_screenshot(f"error_components_page_{product_name}.png")
                        # If we fail, try to refresh and see if it helps before going back.
                        logging.info("Refreshing page to recover from error.")
                        driver.refresh()
                        time.sleep(5)
                        try:
                            logging.info("Attempting to return to main configuration page after error.")
                            back_button = wait.until(EC.element_to_be_clickable((By.ID, "back_list")))
                            driver.execute_script("arguments[0].click();", back_button)
                            time.sleep(3)
                        except Exception as back_e:
                            logging.error(f"Failed to go back to the main page after refresh: {back_e}")
                            raise

            except Exception as e:
                logging.error(f"Failed to edit configuration or add accessories for {product_name}: {e}")
                driver.save_screenshot(f"error_main_product_{product_name}.png")

            logging.info(f"Finished processing product: {product_name}.")

        time.sleep(10) # Final wait to observe the result

    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True)
        driver.save_screenshot('unexpected_error.png')
    finally:
        logging.info("Closing the browser.")
        driver.quit()

if __name__ == "__main__":
    main()