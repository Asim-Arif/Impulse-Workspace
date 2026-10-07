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
        
        const outDir = 'C:\\Users\\Hp\\.gemini\\antigravity\\brain\\85efe67c-31ed-4aff-b07e-6dd6358c8ea4\\scratch\\';
        if (!fs.existsSync(outDir)) {
            fs.mkdirSync(outDir, { recursive: true });
        }

        console.log('Navigating to Register...');
        await page.goto('http://localhost:53535/Account/Register', { waitUntil: 'networkidle2' });
        await page.screenshot({ path: outDir + '01_register.png' });

        console.log('Registering...');
        const email = 'testuser' + Date.now() + '@intraoffice.com';
        await page.type('#register-name', 'Test User');
        await page.type('#register-email', email);
        await page.type('#register-password', 'Test@1234');
        await page.type('#register-confirm', 'Test@1234');
        await page.click('#register-submit');
        
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
        await page.screenshot({ path: outDir + '02_home.png' });

        const routes = [
            { url: '/tasks', name: '03_tasks' },
            { url: '/announcements', name: '04_announcements' },
            { url: '/meetings', name: '05_meetings' },
            { url: '/chat', name: '06_chat' }
        ];

        for (let route of routes) {
            console.log('Navigating to ' + route.url);
            await page.goto('http://localhost:53535' + route.url, { waitUntil: 'networkidle2' });
            await page.screenshot({ path: outDir + route.name + '.png' });
        }

        await browser.close();
        console.log('Done.');
    } catch (err) {
        console.error(err);
    }
})();
