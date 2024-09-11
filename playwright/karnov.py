
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless = False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://pro.karnovgroup.dk/b/areas/karnov-online")

    # Load the cookies
    with open("playwright/cookies.txt", "r") as f:
        cookies = json.loads(f.read())
        context.add_cookies(cookies)
from playwright.sync_api import sync_playwright

import random
import time
def human_like_mouse_move(page, start_x, start_y, end_x, end_y):
    # Simulerer menneskelig mus bevægelse
    steps = 20
    for i in range(steps):
        x = start_x + (end_x - start_x) * (i + 1) / steps
        y = start_y + (end_y - start_y) * (i + 1) / steps
        # Tilføjer lidt tilfældig bevægelse
        x += random.uniform(-5, 5)
        y += random.uniform(-5, 5)
        page.mouse.move(x, y)
        time.sleep(random.uniform(0.05, 0.2))  # Simulerer variation i hastighed

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # Sæt headless=True for at køre uden GUI
    page = browser.new_page()
    page.goto('https://example.com')  # Erstat med URL'en til din testside

    # Simulere musebevægelser
    human_like_mouse_move(page, 100, 100, 400, 400)

    # Her kan du tilføje yderligere interaktioner som klik eller tekstindtastning

    browser.close()
