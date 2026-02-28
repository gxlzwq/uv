import asyncio
from playwright.async_api import async_playwright

async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        try:
            print("Navigating to 12306.cn...")
            await page.goto("https://www.12306.cn/index/", wait_until="networkidle")

            # Station data to set directly
            stations = {
                "from": {"text": "柳州", "code": "LZZ"},
                "to": {"text": "北京西", "code": "BXP"}
            }

            print(f"Setting stations: {stations['from']['text']} -> {stations['to']['text']}...")
            
            # Use evaluate to set values and trigger internal state changes
            await page.evaluate("""(stations) => {
                const setStation = (textId, hiddenId, stationData) => {
                    const textInput = document.getElementById(textId);
                    const hiddenInput = document.getElementById(hiddenId);
                    if (textInput && hiddenInput) {
                        textInput.value = stationData.text;
                        hiddenInput.value = stationData.code;
                        
                        // Dispatch events to notify the page of changes
                        ['input', 'change', 'blur'].forEach(eventName => {
                            textInput.dispatchEvent(new Event(eventName, { bubbles: true }));
                            hiddenInput.dispatchEvent(new Event(eventName, { bubbles: true }));
                        });
                    }
                };

                setStation('fromStationText', 'fromStation', stations.from);
                setStation('toStationText', 'toStation', stations.to);
            }""", stations)

            # 3. Fill in the departure date
            print("Setting departure date to 2026-03-03...")
            await page.locator("#train_date").fill("2026-03-03")

            # 4. Click the "高铁/动车" checkbox
            print("Selecting '高铁/动车'...")
            await page.locator("#isHighDan").click()

            # 5. Click the search button
            print("Clicking search button...")
            await page.locator("#search_one").click()

            print("Search submitted. The browser will stay open for 30 seconds.")
            await asyncio.sleep(30)

        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(run())
