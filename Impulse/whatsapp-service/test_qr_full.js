const puppeteer = require('puppeteer-core');
const fs = require('fs');

(async () => {
    try {
        console.log('Launching browser...');
        const browser = await puppeteer.launch({
            executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        const page = await browser.newPage();
        
        await page.setViewport({ width: 1200, height: 1600 });

        const outDir = 'C:\\Users\\Hp\\.gemini\\antigravity\\brain\\85efe67c-31ed-4aff-b07e-6dd6358c8ea4\\scratch\\';
        
        console.log('Navigating to Register...');
        await page.goto('http://localhost:53535/Account/Register', { waitUntil: 'networkidle2' });

        console.log('Registering...');
        const email = 'testfull3' + Date.now() + '@intraoffice.com';
        await page.type('#register-name', 'Test Full');
        await page.type('#register-email', email);
        await page.type('#register-password', 'Test@1234');
        await page.type('#register-confirm', 'Test@1234');
        await Promise.all([
            page.click('#register-submit'),
            page.waitForNavigation({ waitUntil: 'networkidle2' })
        ]);
        
        console.log('Navigated! Waiting 6 seconds for QR Code...');
        await new Promise(r => setTimeout(r, 6000));
        await page.screenshot({ path: outDir + 'test_home_full.png', fullPage: true });

        await browser.close();
        console.log('Done.');
    } catch (err) {
        console.error(err);
    }
})();
