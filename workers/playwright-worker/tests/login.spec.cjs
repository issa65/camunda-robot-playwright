const { test, expect } = require('@playwright/test');
const { sign } = require('crypto');

test.only('Browser Context Playwright test', async ({ browser }) =>
{
    const context = await browser.newContext();
    const page = await context.newPage();

    await page.goto(
        "https://rahulshettyacademy.com/loginpagePractise/"
    );

    console.log(await page.title());

    const userName = page.locator('#username');
    const signIn = page.locator("#signInBtn");
    const cardTitles = page.locator(".card-body a");

    // css , xpath
    await userName.fill("rahulshetty");

    await page
        .locator("[type='password']")
        .fill("Learning@830$3mK2");

    await page
        .locator("#signInBtn")
        .click();

    // wait until this locator shown up page
    console.log(
        await page
            .locator("[style*='block']")
            .textContent()
    );

    await expect(
        page.locator("[style*='block']")
    ).toContainText('Incorrect');

    // type - fill
    await userName.fill("");

    await userName.fill(
        "rahulshettyacademy"
    );

    await signIn.click();

    // console.log(await cardTitles.first().textContent());
    // console.log(await cardTitles.nth(1).textContent());

    // const allTitles = await cardTitles.allTextContents();

    await page
        .locator(".card-body a")
        .first()
        .waitFor();

    const allTitles =
        await cardTitles.allTextContents();

    console.log(allTitles);
});