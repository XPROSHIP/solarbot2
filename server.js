const express = require('express');
const cors = require('cors');
const axios = require('axios');
const cheerio = require('cheerio');
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

const axiosConfig = {
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' },
    timeout: 15000,
    validateStatus: () => true
};

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

async function fetchHTML(url) {
    try {
        const response = await axios.get(url, axiosConfig);
        return cheerio.load(response.data);
    } catch (e) {
        return null;
    }
}

async function scrapeIdeasoft(site) {
    let siteUrunleri = [];
    let kategoriLinkleri = [];
    
    try {
        const response = await axios.get(site.url, axiosConfig);
        const match = response.data.match(/navigationMenu\s*=\s*(\{.*?\});/);
        if (match) {
            const nav = JSON.parse(match[1]);
            if (nav.categories) {
                nav.categories.forEach(c => { 
                    kategoriLinkleri.push(c.url); 
                    if(c.subCategories) c.subCategories.forEach(sub => kategoriLinkleri.push(sub.url)); 
                });
            }
        } else {
            const $ = cheerio.load(response.data);
            $('a').each((i, el) => {
                let href = $(el).attr('href');
                if (href && (href.includes('kategori') || href.includes('urunler'))) {
                    kategoriLinkleri.push(href);
                }
            });
        }
        
        kategoriLinkleri = [...new Set(kategoriLinkleri)];
        
        for (let url of kategoriLinkleri) {
            let sayfa = 1;
            let devam = true;
            let oncekiVeri = "";
            
            while(devam && sayfa <= 50) {
                try {
                    let tamLink = site.url + url + (url.includes('?') ? '&' : '?') + 'sayfa=' + sayfa;
                    const $ = await fetchHTML(tamLink);
                    if (!$) break;
                    
                    let urunler = [];
                    $('.showcase').each((i, el) => {
                        const iEl = $(el).find('.showcase-title a');
                        const fElNew = $(el).find('.showcase-price-new').text();
                        const fElOld = $(el).find('.showcase-price').text();
                        const fText = fElNew || fElOld;
                        
                        if (iEl.length && fText) {
                            let uLink = iEl.attr('href');
                            if (uLink && !uLink.startsWith('http')) uLink = site.url + uLink;
                            urunler.push({ urun_adi: iEl.text().trim(), fiyat_guncel: fText.trim(), link: uLink });
                        }
                    });
                    
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
    } catch(e) {}
    return siteUrunleri;
}

async function scrapeHeuristic(site) {
    let siteUrunleri = [];
    try {
        const $home = await fetchHTML(site.url);
        if (!$home) return siteUrunleri;
        
        let kategoriLinkleri = [];
        const host = new URL(site.url).hostname;
        
        $home('a').each((i, el) => {
            let href = $home(el).attr('href');
            if (!href) return;
            try {
                let u = new URL(href, site.url);
                let p = u.pathname.toLowerCase();
                if (u.hostname === host && (p.includes('kategori') || p.includes('urunler') || p.split('/').length > 2) && !p.includes('sepet')) {
                    kategoriLinkleri.push(u.href);
                }
            } catch(e) {}
        });
        
        kategoriLinkleri = [...new Set(kategoriLinkleri)];
        
        for (let url of kategoriLinkleri) {
            try {
                const $ = await fetchHTML(url);
                if (!$) continue;
                
                $('span, div, p, b').each((i, el) => {
                    let text = $(el).text().trim();
                    if (/(?=.*\d)(TL|₺|TRY)/i.test(text) && text.length < 30) {
                        let kart = $(el).parent();
                        let limit = 0;
                        let aTag = null;
                        
                        while (kart.length && kart[0].name !== 'body' && limit < 8) {
                            aTag = kart.find('a');
                            if (aTag.length > 0) break;
                            kart = kart.parent();
                            limit++;
                        }
                        
                        if (aTag && aTag.length > 0) {
                            let uAdi = aTag.first().text().trim();
                            let uLink = aTag.first().attr('href');
                            if (uAdi.length > 5 && uLink) {
                                if (!uLink.startsWith('http')) uLink = new URL(uLink, site.url).href;
                                siteUrunleri.push({ urun_adi: uAdi, fiyat_guncel: text, link: uLink });
                            }
                        }
                    }
                });
            } catch (e) {}
        }
    } catch(e){}
    return siteUrunleri;
}

async function scrapeCustom(site) {
    let items = [];
    try {
        const $ = await fetchHTML(site.url);
        if (!$) return items;
        
        $(site.sel.kart).each((i, el) => {
            const iEl = $(el).find(site.sel.isim);
            const fEl = $(el).find(site.sel.fiyat);
            if (iEl.length && fEl.length) {
                let aTag = $(el).find('a').first();
                let link = aTag.attr('href') || site.url;
                if (link && !link.startsWith('http')) link = new URL(link, site.url).href;
                items.push({ urun_adi: iEl.text().trim(), fiyat_guncel: fEl.text().trim(), link: link });
            }
        });
    } catch(e) {}
    return items;
}

app.post('/start', async (req, res) => {
    if (scrapingStatus.isRunning) return res.status(400).json({ error: "Zaten çalışıyor" });
    res.json({ success: true });
    scrapingStatus.isRunning = true;
    
    let globalVeritabani = [];

    try {
        for (let i = 0; i < SITES.length; i++) {
            const site = SITES[i];
            sendSSE({ type: 'info', msg: `Taranıyor: ${site.magaza_adi}`, percent: Math.round(((i + 1) / SITES.length) * 100) });
            
            let hamUrunler = [];
            try {
                if (site.type === "ideasoft") hamUrunler = await scrapeIdeasoft(site);
                else if (site.type === "heuristic") hamUrunler = await scrapeHeuristic(site);
                else if (site.type === "custom") hamUrunler = await scrapeCustom(site);

                hamUrunler.forEach(u => {
                    let temiz = veriTemizle(u.urun_adi, u.link, u.fiyat_guncel);
                    if (temiz) globalVeritabani.push({ magaza: site.magaza_adi, urun_adi: temiz.isim, fiyat_num: temiz.fiyatNum, link: temiz.link });
                });
            } catch (e) {
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
        scrapingStatus.isRunning = false;
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Sunucu http://0.0.0.0:${PORT} adresinde aktif. Dış bağlantılara açık.`);
});