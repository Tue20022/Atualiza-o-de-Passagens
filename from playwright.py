from playwright.sync_api import sync_playwright
import re

def formatar_cidade(valor):
    valor = str(valor).strip().title()
    valor = re.sub(r"\s*-\s*", ", ", valor)
    return valor

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto("https://www.clickbus.com.br/", wait_until="domcontentloaded")
    page.wait_for_timeout(4000)

    try:
        page.locator('button:has-text("Aceitar")').click(timeout=2000)
    except:
        pass

    origem = formatar_cidade("ALEGRE-ES")

    campo = page.locator("#origin")
    campo.click()
    page.wait_for_timeout(300)
    campo.fill("")
    page.wait_for_timeout(300)
    campo.type(origem, delay=100)
    page.wait_for_timeout(3000)

    seletores = [
        '[role="listbox"] [role="option"]',
        '[role="listbox"] li',
        'ul li',
        '[class*="autocomplete"] li',
        '[class*="suggest"] li',
        '[class*="menu"] li',
        '[class*="dropdown"] li',
    ]

    achou = False

    for sel in seletores:
        try:
            loc = page.locator(sel)
            total = loc.count()
            if total == 0:
                continue

            print(f"\nSELETOR: {sel}")
            for i in range(min(total, 15)):
                try:
                    item = loc.nth(i)
                    if item.is_visible():
                        txt = item.inner_text(timeout=1000).strip()
                        if txt:
                            print(f"[{i}] {repr(txt)}")
                            achou = True
                except:
                    pass
        except:
            pass

    if not achou:
        print("Nenhuma opção visível encontrada.")

    print("origem visível:", repr(page.locator('#origin').input_value()))
    print("origem final:", repr(page.locator('#origin-final').get_attribute('value')))

    input("Enter para fechar...")
    browser.close()