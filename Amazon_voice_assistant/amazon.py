import speech_recognition as sr
import pyttsx3
import time
import csv
import os
import re
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime

recognizer = sr.Recognizer()
mic = sr.Microphone()
engine = pyttsx3.init()

def speak(text):
    print(">>", text)
    engine.say(text)
    engine.runAndWait()

def listen(prompt):
    while True:
        speak(prompt)
        with mic as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            print("Listening...")
            try:
                audio = recognizer.listen(source, timeout=10, phrase_time_limit=10)
                text = recognizer.recognize_google(audio).strip().lower()
                return text
            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that. Please repeat.")
            except sr.WaitTimeoutError:
                speak("No voice input detected. Please try again.")
            except sr.RequestError:
                speak("There was a problem with the speech service.")
                return ""

def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email)

def get_spelled_email():
    while True:
        spoken = listen("Please spell your email address. For example, say: n a m e at gmail dot com")
        spoken = spoken.replace(" at ", "@").replace(" dot ", ".")
        spoken = spoken.replace("underscore", "_").replace("dash", "-").replace(" ", "")
        print(f"Processed Email: {spoken}")
        if is_valid_email(spoken):
            speak(f"You entered email as: {spoken}")
            return spoken
        else:
            speak("That doesn't look like a valid email address. Let's try again.")


def get_spoken_password():
    spoken = listen("Please say your Amazon password.")
    spoken = spoken.replace(" at ", "@").replace(" dot ", ".")
    spoken = spoken.replace("underscore", "_").replace("dash", "-")
    spoken = spoken.replace(" space ", "").replace(" ", "")

    # Print and speak password for verification
    speak("This is your password. Please verify it on the screen.")
    sys.stdout.write(f"\nYour password is: {spoken}")
    sys.stdout.flush()
    time.sleep(3)
    sys.stdout.write(f"\rYour password is: {'*' * len(spoken)}\n")
    sys.stdout.flush()

    speak("Password verification complete.")
    return spoken


def save_to_csv(data):
    filename = f"amazon_viewed_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    filepath = os.path.join(os.getcwd(), filename)

    with open(filepath, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=[
            "ASIN", "Title", "Price", "Image_URL", "Date", "Time", "Added_to_Cart", "Amazon_Link"
        ])
        writer.writeheader()
        writer.writerows(data)

    speak(f"All product details have been saved to the file named {filename}")
    print(f">> Saved to {filename}")

# -------------------- LOGIN FLOW --------------------
email = get_spelled_email()
password = get_spoken_password()

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://www.amazon.in/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.in%2F%3Ftag%3Dmsndeskabkin-21&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=inflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0")

try:
    wait = WebDriverWait(driver, 15)

    email_input = wait.until(EC.presence_of_element_located((By.ID, "ap_email")))
    email_input.send_keys(email)
    driver.find_element(By.ID, "continue").click()

    password_input = wait.until(EC.presence_of_element_located((By.ID, "ap_password")))
    password_input.send_keys(password)
    driver.find_element(By.ID, "signInSubmit").click()

    speak("Please complete CAPTCHA or OTP if prompted.")
    input(">> Press ENTER after completing CAPTCHA or OTP to continue...")

except Exception as e:
    speak("An error occurred during login.")
    print("Error:", e)

speak("Automation will now resume.")

# -------------------- MAIN LOOP --------------------
all_products = []

while True:
    product_name = listen("What product would you like to search for on Amazon? Or say 'exit' to stop.")
    if "exit" in product_name or "quit" in product_name or "close" in product_name:
        speak("Okay, ending the Amazon automation. Thank you!")
        break

    if not product_name.strip():
        speak("No valid product name heard. Let's try again.")
        continue

    search_query = product_name.replace(" ", "+")
    search_url = f"https://www.amazon.in/s?k={search_query}"

    speak(f"Searching Amazon for {product_name}")
    driver.get(search_url)

    speak("Search complete. Please use your mouse to select a product to view.")
    input(">> Press ENTER after you've selected a product to continue...")

    time.sleep(3)
    windows = driver.window_handles
    if len(windows) > 1:
        driver.switch_to.window(windows[1])
        speak("Switched to the selected product tab.")
    else:
        speak("Product did not open in a new tab. Staying on current page.")

    # Extract product info
    try:
        title = driver.find_element(By.ID, "productTitle").text.strip()
    except:
        title = "N/A"

    price = "N/A"
    price_selectors = [
        "#priceblock_ourprice",
        "#priceblock_dealprice",
        "#priceblock_saleprice",
        "#corePriceDisplay_desktop_feature_div .a-price .a-offscreen",
        "#corePrice_feature_div .a-price .a-offscreen"
    ]

    # Try all known selectors
    for selector in price_selectors:
        try:
            element = driver.find_element(By.CSS_SELECTOR, selector)
            if element and element.text.strip():
                price = element.text.strip()
                break
        except:
            continue

    # Fallback: Combine a-price-whole and a-price-fraction
    if price == "N/A":
        try:
            whole = driver.find_element(By.CLASS_NAME, "a-price-whole").text.strip().replace(",", "")
            fraction = driver.find_element(By.CLASS_NAME, "a-price-fraction").text.strip()
            price = f"₹{whole}.{fraction}"
        except:
            pass
    try:
        img_url = driver.find_element(By.ID, "landingImage").get_attribute("src")
    except:
        img_url = "N/A"

    try:
        asin = driver.find_element(By.XPATH, "//th[text()='ASIN']/following-sibling::td").text.strip()
    except:
        asin = driver.current_url.split("/dp/")[1].split("/")[0] if "/dp/" in driver.current_url else "N/A"

    product_link = driver.current_url
    now = datetime.now()
    product_info = {
        "ASIN": asin,
        "Title": title,
        "Price": price,
        "Image_URL": img_url,
        "Date": now.strftime("%Y-%m-%d"),
        "Time": now.strftime("%H:%M:%S"),
        "Added_to_Cart": "No",
        "Amazon_Link": product_link
    }
    # Speak out product info
    speak("Here is the product information I found.")
    if title != "N/A":
        speak(f"The title is: {title}")
    else:
        speak("I couldn't find the product title.")

    if price != "N/A":
        speak(f"The price is: {price}")
    else:
        speak("I couldn't find the price.")

    # Add to cart prompt
    decision = listen("Would you like to add this product to your cart? Please say yes or no.")
    decision = decision.replace(" ", "").replace("-", "").lower()

    if decision in ["yes", "y", "s", "yyes", "ys", "yees"]:
        product_info["Added_to_Cart"] = "Yes"
        speak("Trying to add the product to your cart...")

        try:
            cart_selectors = [
                (By.ID, "add-to-cart-button"),
                (By.ID, "submit.add-to-cart"),
                (By.ID, "submit.add-to-cart-ubb"),
                (By.CSS_SELECTOR, "input#add-to-cart-button"),
                (By.CSS_SELECTOR, "input[name='submit.add-to-cart']"),
                (By.XPATH, "//input[@value='Add to Cart']"),
                (By.XPATH, "//button[contains(text(), 'Add to Cart')]")
            ]

            button_clicked = False
            for selector in cart_selectors:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located(selector))
                    btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(selector))

                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    driver.execute_script("arguments[0].style.border='3px solid red'", btn)
                    time.sleep(1)

                    try:
                        btn.click()
                    except:
                        driver.execute_script("arguments[0].click();", btn)

                    speak(f"Clicked Add to Cart using selector: {selector}")
                    button_clicked = True
                    break
                except Exception as e:
                    print(f">> Failed selector {selector}: {e}")

            if not button_clicked:
                speak("Could not find or click any known Add to Cart button.")
            else:
                try:
                    WebDriverWait(driver, 6).until(
                        EC.presence_of_element_located((By.ID, "attachDisplayAddBaseAlert"))
                    )
                    speak("Product added to your cart successfully!")
                except:
                    speak("Tried adding to cart. Confirmation not found but it might still be added.")
        except Exception as e:
            speak("Something went wrong while trying to add the product to your cart.")
            print(">> Add to cart final error:", e)
    else:
        speak("Product not added to cart.")

    # Final cart confirmation
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "sw-gtc"))
        )
        print("Product was added to the cart.")
    except:
        print("Product may not have been added to the cart. Please check manually.")

    all_products.append(product_info)
    # Clean up: Close product tab if it was opened in a new window
    if len(driver.window_handles) > 1:
        driver.close()  # Closes current tab
        driver.switch_to.window(driver.window_handles[0])  # Go back to main tab

    # Navigate back to main Amazon homepage
    driver.get("https://www.amazon.in/")
    again = listen("Would you like to search for another product?")
    again = again.replace(" ", "").lower()
    if again not in ["yes", "y", "yeah", "sure", "yea", "yep"]:
        speak("Okay, ending the session.")
        break

save_to_csv(all_products)
driver.quit()