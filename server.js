const express = require('express');
const cors = require('cors');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

const SITES = [
    { magaza_adi: "Kamu Solar", url: "https://www.kamusolar.com", type: "ideasoft" },
    { magaza_adi: "Global Enerji", url: "https://www.globalenerjimarketim.com", type: "ideasoft" },
    { magaza_adi: "Enerji Pazarı", url: "https://www.enerjipazari.com.tr", type: "ideasoft" },
    { magaza_adi: "Yapı Bahçe", url: "https://www.yapibahce.com", type: "ideasoft" },
    { magaza_adi: "Teknovasyon Arge", url: "https://www.teknovasyonarge.com", type: "ideasoft" },
    { magaza_adi: "Solar Sanal Market", url: "https://www.solarsanalmarket.com", type: "ideasoft" },
    { magaza_adi: "Kampa", url: "https://www.kampa.com.tr", type: "ideasoft" },
    { magaza_adi: "Nonstop Enerji", url: "https://www.nonstopenerji.com", type: "ideasoft" },
    { magaza_adi: "Gümüş Solar", url: "https://www.gumussolar.com", type: "ideasoft" },
    { magaza_adi: "Alize Marin Market", url: "https://www.alizemarinmarket.com", type: "ideasoft" },
    { magaza_adi: "Solenser Market", url: "https://www.solensermarket.com", type: "ideasoft" },
    { magaza_adi: "Solar İst Shop", url: "https://www.solaristshop.com", type: "ideasoft" },
    { magaza_adi: "Modül Elektronik", url: "https://www.modulelektronik.com", type: "ideasoft" },
    { magaza_adi: "Urla Solar", url: "https://urlasolar.com", type: "heuristic" },
    { magaza_adi: "Sakarya Solar", url: "https://sakaryasolarmarket.com", type: "heuristic" },
    { magaza_adi: "Solar Zirve", url: "https://www.solarzirve.com", type: "heuristic" },
    { magaza_adi: "Tam Solar", url: "https://tamsolar.com.tr", type: "custom", sel: { kart: '.product-container', isim: '.product-name', fiyat: '.sell-price' } },
    { magaza_adi: "Atakale", url: "https://www.atakale.com.tr", type: "custom", sel: { kart: '.product-thumb', isim: '.caption h4 a', fiyat: '.price' } },
    { magaza_adi: "Enerjimar", url: "https://enerjimar.com", type: "custom", sel: { kart: '.urun-kutusu', isim: 'h2 a', fiyat: '.urun-fiyat' } },
    { magaza_adi: "İda Solar", url: "https://www.idasolar.com", type: "custom", sel: { kart: '.card-product', isim: '.title', fiyat: '.sale-price' } }
];

let scrapingStatus = { isRunning: false, clients: [] };

function sendSSE(data) {
    scrapingStatus.clients.forEach(res => res.write(`data: ${JSON.stringify(data)}\n\n`));
}

app.get('/progress', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();
    scrapingStatus.clients.push(res);
    req.on('close', () => { scrapingStatus.clients = scrapingStatus.clients.filter(c => c !== res); });
});

function veriTemizle(isim, link, fiyatText) {
    if (!isim || !link || !fiyatText) return null;
    let upIsim = isim.toUpperCase().trim();
    if (upIsim === 'TÜKENDİ' || upIsim === 'PAYLAŞ' || upIsim.length < 5) return null;
    let temizFiyat = fiyatText.replace(/tl|₺|try|lira|kdv|vergi|dahil|indirimli|\+|/gi, '').trim();
    let fiyatParcalari = temizFiyat.split(/\s+/);
    temizFiyat = fiyatParcalari[fiyatParcalari.length - 1];
    let numStr = temizFiyat.replace(/[^0-9,.]/g, '');
    if (!numStr || numStr === "0") return null;
    let floatVal = 0;
    let sonVirgul = numStr.lastIndexOf(',');
    let sonNokta = numStr.lastIndexOf('.');
    if (sonVirgul > sonNokta) floatVal = parseFloat(numStr.replace(/\./g, '').replace(',', '.'));
    else if (sonNokta > sonVirgul) floatVal = parseFloat(numStr.replace(/,/g, ''));
    else floatVal = parseFloat(numStr);
    if (isNaN(floatVal) || floatVal <= 0) return null;
    return { isim: isim.replace(/\n/g, ' ').trim(), fiyatNum: floatVal, link: link };
}

async function scrapeIdeasoft(page, site) {
    let siteUrunleri = [];
    try {
        const kategoriLinkleri = await page.evaluate(() => {
            let l = [];
            if (typeof navigationMenu !== 'undefined' && navigationMenu.categories) {
                navigationMenu.categories.forEach(a => { l.push(a.url); if(a.subCategories) a.subCategories.forEach(sub => l.push(sub.url)); });
            }
            return l; 
        });
        
        for (let url of kategoriLinkleri) {
            let sayfa = 1;
            let devam = true;
            let oncekiVeri = "";
            
            while(devam && sayfa <= 15) { 
                try {
                    let link = site.url + url + (url.includes('?') ? '&' : '?') + 'sayfa=' + sayfa;
                    await page.goto(link, { waitUntil: 'domcontentloaded', timeout: 30000 });
                    
                    const urunler = await page.evaluate((b) => {
                        let toplanan = [];
                        document.querySelectorAll('.showcase').forEach(el => {
                            const iEl = el.querySelector('.showcase-title a');
                            let fEl = el.querySelector('.showcase-price-new')?.innerText || el.querySelector('.showcase-price')?.innerText;
                            if (iEl && fEl) toplanan.push({ urun_adi: iEl.innerText, fiyat_guncel: fEl, link: iEl.href.startsWith('http') ? iEl.href : b + iEl.getAttribute('href') });
                        });
                        return toplanan;
                    }, site.url);
                    
                    let guncelVeri = JSON.stringify(urunler);
                    if(urunler.length === 0 || guncelVeri === oncekiVeri) {
                        devam = false;
                    } else {
                        siteUrunleri = siteUrunleri.concat(urunler);
                        oncekiVeri = guncelVeri;
                        sayfa++;
                    }
                } catch (e) { devam = false; }
            }
        }
    } catch(e){}
    return siteUrunleri;
}

async function scrapeHeuristic(page, site) {
    let siteUrunleri = [];
    try {
        const kategoriLinkleri = await page.evaluate(() => {
            const host = window.location.hostname;
            return [...new Set(Array.from(document.querySelectorAll('a')).map(a => a.href).filter(h => {
                try { const u = new URL(h); const p = u.pathname.toLowerCase(); return u.hostname === host && (p.includes('kategori') || p.includes('urunler') || p.split('/').length > 2) && !p.includes('sepet'); } catch(e){ return false; }
            }))];
        });
        for (let url of kategoriLinkleri) {
            try {
                await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
                const urunler = await page.evaluate(() => {
                    let toplanan = [];
                    Array.from(document.querySelectorAll('span, div, p, b')).forEach(el => {
                        const t = el.innerText.trim();
                        if (/(?=.*\d)(TL|₺|TRY)/i.test(t) && t.length < 30) {
                            let k = el.parentElement; let limit = 0;
                            while (k && k.tagName !== 'BODY' && limit < 8) { if (k.querySelector('a')) break; k = k.parentElement; limit++; }
                            if (k && k.tagName !== 'BODY') {
                                const a = k.querySelector('a');
                                if (a && a.innerText.trim().length > 5) toplanan.push({ urun_adi: a.innerText.trim(), fiyat_guncel: t, link: a.href });
                            }
                        }
                    });
                    return toplanan;
                });
                siteUrunleri = siteUrunleri.concat(urunler);
            } catch (e) { }
        }
    } catch(e){}
    return siteUrunleri;
}

async function scrapeCustom(page, site) {
    try {
        return await page.evaluate((s, url) => {
            let items = [];
            document.querySelectorAll(s.kart).forEach(el => {
                const iEl = el.querySelector(s.isim); const fEl = el.querySelector(s.fiyat);
                if (iEl && fEl) items.push({ urun_adi: iEl.innerText, fiyat_guncel: fEl.innerText, link: el.querySelector('a')?.href || url });
            });
            return items;
        }, site.sel, site.url);
    } catch(e) { return []; }
}

app.post('/start', async (req, res) => {
    if (scrapingStatus.isRunning) return res.status(400).json({ error: "Zaten çalışıyor" });
    res.json({ success: true });
    scrapingStatus.isRunning = true;
    
    let globalVeritabani = [];
    const browser = await puppeteer.launch({ 
        headless: "new", 
        args: [
            '--no-sandbox', 
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--single-process',
            '--no-zygote'
        ] 
    });

    try {
        for (let i = 0; i < SITES.length; i++) {
            const site = SITES[i];
            sendSSE({ type: 'info', msg: `Taranıyor: ${site.magaza_adi}`, percent: Math.round(((i + 1) / SITES.length) * 100) });
            
            const page = await browser.newPage();
            await page.setRequestInterception(true);
            page.on('request', (request) => {
                if (['image', 'stylesheet', 'font', 'media'].includes(request.resourceType())) {
                    request.abort();
                } else {
                    request.continue();
                }
            });
            await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64)');
            
            let hamUrunler = [];
            try {
                await page.goto(site.url, { waitUntil: 'domcontentloaded', timeout: 60000 });
                if (site.type === "ideasoft") hamUrunler = await scrapeIdeasoft(page, site);
                else if (site.type === "heuristic") hamUrunler = await scrapeHeuristic(page, site);
                else if (site.type === "custom") hamUrunler = await scrapeCustom(page, site);

                hamUrunler.forEach(u => {
                    let temiz = veriTemizle(u.urun_adi, u.link, u.fiyat_guncel);
                    if (temiz) globalVeritabani.push({ magaza: site.magaza_adi, urun_adi: temiz.isim, fiyat_num: temiz.fiyatNum, link: temiz.link });
                });
            } catch (e) {
            } finally {
                await page.close();
            }
        }

        globalVeritabani = [...new Map(globalVeritabani.map(item => [item.link, item])).values()];

        fs.writeFileSync(path.join(__dirname, 'KamuSolar_Veritabani.json'), JSON.stringify(globalVeritabani, null, 2), 'utf-8');

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Guncel Fiyatlar');
        sheet.columns = [
            { header: 'Mağaza', key: 'magaza', width: 25 },
            { header: 'Ürün Adı', key: 'urun_adi', width: 80 },
            { header: 'Fiyat', key: 'fiyat_num', width: 15 },
            { header: 'Link', key: 'link', width: 40 }
        ];
        globalVeritabani.forEach(u => sheet.addRow(u));
        await workbook.xlsx.writeFile(path.join(__dirname, 'KamuSolar_Rapor.xlsx'));

        sendSSE({ type: 'done', percent: 100, jsonUrl: '/KamuSolar_Veritabani.json', excelUrl: '/KamuSolar_Rapor.xlsx' });
    } catch (err) {
        sendSSE({ type: 'info', msg: "Hata oluştu", percent: 0 });
    } finally {
        await browser.close();
        scrapingStatus.isRunning = false;
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Sunucu http://0.0.0.0:${PORT} adresinde aktif. Dış bağlantılara açık.`);
});