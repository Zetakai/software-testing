# -*- coding: utf-8 -*-
"""Live execution of SauceDemo test cases via Selenium headless."""
import time, json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

URL = "https://www.saucedemo.com/"
PWD = "secret_sauce"
results = []


def driver():
    o = Options()
    o.add_argument("--headless=new")
    o.add_argument("--no-sandbox")
    o.add_argument("--disable-dev-shm-usage")
    o.add_argument("--window-size=1280,1024")
    d = webdriver.Chrome(options=o)
    d.set_page_load_timeout(40)
    return d


def login(d, user):
    d.get(URL)
    WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, "user-name")))
    d.find_element(By.ID, "user-name").clear()
    d.find_element(By.ID, "user-name").send_keys(user)
    d.find_element(By.ID, "password").clear()
    d.find_element(By.ID, "password").send_keys(PWD)
    d.find_element(By.ID, "login-button").click()
    time.sleep(1.5)


def rec(tc, desc, status, actual):
    results.append({"tc": tc, "desc": desc, "status": status, "actual": actual})
    print(f"{tc:8} {status:8} {actual}")



def jsclick(d, css):
    el = WebDriverWait(d, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))
    d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    d.execute_script("arguments[0].click();", el)
    return el


def react_set(d, css, value):
    """Set a React-controlled input so its state updates reliably."""
    el = WebDriverWait(d, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))
    d.execute_script("""
        const el = arguments[0], v = arguments[1];
        const proto = window.HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
        setter.call(el, v);
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, el, value)
    return el


def click_until(d, css, pred, tries=6):
    """Headless clicks on this site are flaky; retry native+JS until pred() holds."""
    import time as _t
    for _ in range(tries):
        try:
            el = WebDriverWait(d, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))
            d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            try:
                el.click()
            except Exception:
                pass
            _t.sleep(0.5)
            if pred():
                return True
            d.execute_script("arguments[0].click();", el)
            _t.sleep(0.7)
            if pred():
                return True
        except Exception:
            pass
        _t.sleep(0.5)
    return pred()

def safe(fn):
    try:
        fn()
    except Exception as e:
        rec("ERR", fn.__name__, "ERROR", f"{type(e).__name__}: {str(e)[:120]}")


# TC-01 valid login
def tc01():
    d = driver()
    try:
        login(d, "standard_user")
        ok = "/inventory.html" in d.current_url
        rec("TC-01", "Login valid standard_user", "PASS" if ok else "FAIL",
            f"url={d.current_url}")
    finally:
        d.quit()


# TC-02 invalid login
def tc02():
    d = driver()
    try:
        d.get(URL)
        WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, "user-name")))
        d.find_element(By.ID, "user-name").send_keys("wrong_user")
        d.find_element(By.ID, "password").send_keys("wrong_pw")
        d.find_element(By.ID, "login-button").click()
        time.sleep(1)
        msg = d.find_element(By.CSS_SELECTOR, "[data-test='error']").text
        ok = "do not match" in msg.lower()
        rec("TC-02", "Login invalid", "PASS" if ok else "FAIL", msg)
    finally:
        d.quit()


# TC-03 empty fields
def tc03():
    d = driver()
    try:
        d.get(URL)
        WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, "login-button")))
        d.find_element(By.ID, "login-button").click()
        time.sleep(1)
        msg = d.find_element(By.CSS_SELECTOR, "[data-test='error']").text
        ok = "username is required" in msg.lower()
        rec("TC-03", "Login field kosong", "PASS" if ok else "FAIL", msg)
    finally:
        d.quit()


# TC-04 locked out
def tc04():
    d = driver()
    try:
        login(d, "locked_out_user")
        msg = d.find_element(By.CSS_SELECTOR, "[data-test='error']").text
        ok = "locked out" in msg.lower()
        rec("TC-04", "Login user terkunci", "PASS" if ok else "FAIL", msg)
    finally:
        d.quit()


# TC-05 add to cart badge
def tc05():
    d = driver()
    try:
        login(d, "standard_user")
        jsclick(d, "button[data-test^='add-to-cart']")
        time.sleep(0.8)
        badge = d.find_element(By.CLASS_NAME, "shopping_cart_badge").text
        ok = badge == "1"
        rec("TC-05", "Add to cart badge", "PASS" if ok else "FAIL", f"badge={badge}")
    finally:
        d.quit()


# TC-06 sort price low->high (standard)
def tc06():
    d = driver()
    try:
        login(d, "standard_user")
        Select(d.find_element(By.CLASS_NAME, "product_sort_container")).select_by_value("lohi")
        time.sleep(1)
        prices = [float(e.text.replace("$", "")) for e in
                  d.find_elements(By.CLASS_NAME, "inventory_item_price")]
        ok = prices == sorted(prices)
        rec("TC-06", "Sort price low-high (standard)", "PASS" if ok else "FAIL",
            f"prices={prices}")
    finally:
        d.quit()


# TC-07 checkout end-to-end (standard)
def tc07():
    d = driver()
    try:
        login(d, "standard_user")
        has = lambda frag: frag in d.current_url
        click_until(d, "button[data-test^='add-to-cart']",
                    lambda: bool(d.find_elements(By.CLASS_NAME, "shopping_cart_badge")))
        click_until(d, ".shopping_cart_link", lambda: has("cart.html"))
        click_until(d, "#checkout", lambda: has("checkout-step-one"))
        react_set(d, "#first-name", "Farid")
        react_set(d, "#last-name", "Zaki")
        react_set(d, "#postal-code", "12345")
        click_until(d, "#continue", lambda: has("checkout-step-two"))
        click_until(d, "#finish", lambda: has("checkout-complete"))
        msg = WebDriverWait(d, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "complete-header"))).text
        ok = "thank you" in msg.lower()
        rec("TC-07", "Checkout end-to-end (standard)", "PASS" if ok else "FAIL", msg)
    finally:
        d.quit()


# TC-08 problem_user images identical
def tc08():
    d = driver()
    try:
        login(d, "problem_user")
        srcs = [e.get_attribute("src") for e in
                d.find_elements(By.CLASS_NAME, "inventory_item_img")]
        uniq = set(srcs)
        # PASS = images correct (many unique). FAIL = all identical.
        if len(uniq) <= 2:
            rec("TC-08", "problem_user images unique", "FAIL",
                f"{len(srcs)} imgs, {len(uniq)} unique -> all identical (DEF-01)")
        else:
            rec("TC-08", "problem_user images unique", "PASS",
                f"{len(srcs)} imgs, {len(uniq)} unique")
    finally:
        d.quit()


# TC-09 error_user checkout last name
def tc09():
    d = driver()
    try:
        login(d, "error_user")
        jsclick(d, "button[data-test^='add-to-cart']")
        time.sleep(0.5)
        jsclick(d, ".shopping_cart_link")
        time.sleep(0.8)
        jsclick(d, "#checkout")
        time.sleep(0.8)
        d.find_element(By.ID, "first-name").send_keys("Farid")
        ln = d.find_element(By.ID, "last-name")
        ln.send_keys("Zaki")
        d.find_element(By.ID, "postal-code").send_keys("12345")
        typed = ln.get_attribute("value")
        d.find_element(By.ID, "continue").click()
        time.sleep(0.8)
        url = d.current_url
        proceeded = "checkout-step-two" in url
        # PASS = last name accepted AND proceeded
        if typed == "Zaki" and proceeded:
            rec("TC-09", "error_user checkout last name", "PASS",
                f"lastname='{typed}', proceeded={proceeded}")
        else:
            rec("TC-09", "error_user checkout last name", "FAIL",
                f"lastname='{typed}', proceeded={proceeded}, url={url} (DEF-02)")
    finally:
        d.quit()


# TC-06b problem_user sorting
def tc06b():
    d = driver()
    try:
        login(d, "problem_user")
        before = [e.text for e in d.find_elements(By.CLASS_NAME, "inventory_item_name")]
        try:
            Select(d.find_element(By.CLASS_NAME, "product_sort_container")).select_by_value("za")
        except Exception as e:
            rec("TC-06b", "problem_user sorting", "FAIL",
                f"sort dropdown error: {type(e).__name__} (DEF-03)")
            return
        time.sleep(1)
        after = [e.text for e in d.find_elements(By.CLASS_NAME, "inventory_item_name")]
        expected = sorted(before, reverse=True)
        if after == expected:
            rec("TC-06b", "problem_user sorting", "PASS", "Z-A order correct")
        else:
            rec("TC-06b", "problem_user sorting", "FAIL",
                f"order unchanged/wrong (DEF-03). after[0]={after[0] if after else None}")
    finally:
        d.quit()


# TC-05b problem_user remove
def tc05b():
    d = driver()
    try:
        login(d, "problem_user")
        jsclick(d, "button[data-test^='add-to-cart']")
        time.sleep(0.8)
        try:
            badge1 = d.find_element(By.CLASS_NAME, "shopping_cart_badge").text
        except Exception:
            badge1 = "0"
        # try click a Remove button
        try:
            jsclick(d, "button[data-test^='remove']")
            time.sleep(0.8)
            try:
                badge2 = d.find_element(By.CLASS_NAME, "shopping_cart_badge").text
            except Exception:
                badge2 = "0"
            if badge2 == "0" or badge2 != badge1:
                rec("TC-05b", "problem_user remove", "PASS",
                    f"badge {badge1}->{badge2}")
            else:
                rec("TC-05b", "problem_user remove", "FAIL",
                    f"badge stuck at {badge2} after Remove (DEF-04)")
        except Exception as e:
            rec("TC-05b", "problem_user remove", "FAIL",
                f"no working Remove btn: {type(e).__name__} (DEF-04)")
    finally:
        d.quit()


# TC-10 logout
def tc10():
    d = driver()
    try:
        login(d, "standard_user")
        click_until(d, "#react-burger-menu-btn",
                    lambda: bool(d.find_elements(By.ID, "logout_sidebar_link")))
        click_until(d, "#logout_sidebar_link",
                    lambda: bool(d.find_elements(By.ID, "login-button")))
        ok = bool(d.find_elements(By.ID, "login-button")) and "inventory" not in d.current_url
        rec("TC-10", "Logout", "PASS" if ok else "FAIL", f"url={d.current_url}")
    finally:
        d.quit()


for fn in [tc01, tc02, tc03, tc04, tc05, tc06, tc07, tc10, tc08, tc09, tc06b, tc05b]:
    safe(fn)

with open("live_results.json", "w") as f:
    json.dump(results, f, indent=2)
p = sum(1 for r in results if r["status"] == "PASS")
fl = sum(1 for r in results if r["status"] == "FAIL")
print(f"\nSUMMARY: {len(results)} run, {p} PASS, {fl} FAIL")
