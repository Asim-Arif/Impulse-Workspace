const puppeteer = require('puppeteer-core');
const fs = require('fs');

(async () => {
    try {
        const browser = await puppeteer.launch({
            executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        const page = await browser.newPage();
        
        page.on('console', msg => console.log('BROWSER LOG:', msg.text()));
        page.on('pageerror', err => console.log('BROWSER ERROR:', err.toString()));
        
        await page.goto('http://localhost:53535/Account/Login', { waitUntil: 'networkidle2' });

        const email = 'testfull' + Date.now() + '@intraoffice.com';
        await page.goto('http://localhost:53535/Account/Register', { waitUntil: 'networkidle2' });

        await page.type('#register-name', 'Test Full');
        await page.type('#register-email', email);
        await page.type('#register-password', 'Test@1234');
        await page.type('#register-confirm', 'Test@1234');
        await Promise.all([
            page.click('#register-submit'),
            page.waitForNavigation({ waitUntil: 'networkidle2' })
        ]);
        
        await new Promise(r => setTimeout(r, 6000));
        await page.screenshot({ path: 'C:\\Users\\Hp\\.gemini\\antigravity\\brain\\85efe67c-31ed-4aff-b07e-6dd6358c8ea4\\scratch\\test_qr_final.png', fullPage: true });
        await browser.close();
        console.log('Done.');
    } catch (err) {
        console.error(err);
    }
})();
