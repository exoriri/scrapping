const { default: axios } = require('axios');
const { constants } = require('buffer');
const cheerio = require('cheerio');
const Excel = require('exceljs');
const fs = require('fs');
const puppeteer = require('puppeteer');

const workbook = new Excel.Workbook();
const workbookWriting = new Excel.Workbook();

const fetchPhones = async (website) => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(website);
    await page.waitForSelector('body');
    const content = await page.content();
    await browser.close();

    const $ = cheerio.load(content);
    const telLinks = $('a[href^="tel:"]');

    let phones = "";

    for (let i = 0; i < telLinks.length; i++) {
        phones += `${telLinks[i].attribs['href'].replace('tel:', '')}, `;
    };

    return phones;
};

const getScrappedData = async (website, protocol) => {
    const tel = await fetchPhones(`${protocol}://${website}/`);
    return tel;
};

(async () => (await workbook.xlsx.readFile('semrush.cont.xlsx')))()
    .then(async xlsxFile => {
        const res = {};
        const worksheet = xlsxFile.getWorksheet();
        const ranks = worksheet.getColumn(1).values;
        const websites = worksheet.getColumn(2).values;
        const organicKeywords = worksheet.getColumn(3).values;
        const organicTraffics = worksheet.getColumn(4).values;
        const organicCost = worksheet.getColumn(5).values;
        const adwordsKeywords = worksheet.getColumn(6).values;
        const adwordsTraffics = worksheet.getColumn(7).values;
        const adwordsCosts = worksheet.getColumn(8).values;
        const phoneNumbers = worksheet.getColumn(9).values;
        const additionalInfo = worksheet.getColumn(10).values;
        const names = worksheet.getColumn(11).values;
        const statuses = worksheet.getColumn(12).values;

        const contactsWorksheet = workbookWriting.addWorksheet('Лист 1');

        contactsWorksheet.columns = [
            { key: "rank", header: "Rank" },
            { key: "domain", header: "Domain" },
            { key: "organicKeywords", header: "Organic Keywords" },
            { key: "organicTraffic", header: "Organic Traffic" },
            { key: "organicCost", header: "Organic Cost" },
            { key: "adwordsKeywords", header: "Adwords Keywords" },
            { key: "adwordsTraffic", header: "Adwords Traffic" },
            { key: "adwordsCost", header: "Adwords Cost" },
            { key: "phoneNumber", header: "Phone number" },
            { key: "anotherContact", header: "Another contact" },
            { key: "name", header: "Name" },
            { key: "comment", header: "Comment" },
            { key: "status", header: "Status" }
        ];

        const indexToAdd = 2;
        let count = 0;

        for (let i = indexToAdd; i < websites.length; i++) {
            const contacts = {
                tel: {error : '#N/A'},
                vk: {error : '#N/A'}
            };

            if (!phoneNumbers[i].error && phoneNumbers[i] !== '#N/A') {
                count += 1;
            }
            
            if (phoneNumbers[i].error || phoneNumbers[i] === '#N/A') {
                try {
                    const tel = await getScrappedData(websites[i], 'https');
                    if (tel.length === 0) {
                        const tel = await getScrappedData(`${websites[i]}/contacts`, 'https');
                        if (tel.length === 0) {
                            const tel = await getScrappedData(`${websites[i]}/contact`, 'https');
                            contacts.tel = tel;
                        } else {
                            contacts.tel = tel;
                        }
                    } else {
                        contacts.tel = tel;
                    };
                } catch(e) {
                    console.log('ERROR', e.message);
                }
            } else {
                contacts.tel = phoneNumbers[i];
            }

            res[websites[i]] = contacts;

            console.log('RESULT', res);

        };

        Object.keys(res).forEach((key, i) => {
            const phoneNumber = res[key].tel ? res[key].tel : { error: '#N/A' };
            contactsWorksheet.addRow({
                rank: ranks[i+indexToAdd],
                domain: websites[i+indexToAdd],
                organicKeywords: organicKeywords[i+indexToAdd],
                organicTraffic: organicTraffics[i+indexToAdd],
                organicCost: organicCost[i+indexToAdd],
                adwordsKeywords: adwordsKeywords[i+indexToAdd],
                adwordsTraffic: adwordsTraffics[i+indexToAdd],
                adwordsCost: adwordsCosts[i+indexToAdd],
                phoneNumber,
                anotherContact: additionalInfo[i+indexToAdd],
                name: names[i+indexToAdd],
                comment: '',
                status: statuses[i+indexToAdd]
            });
        })

        const contactsXlsx = await workbookWriting.xlsx.writeFile('contacts.xlsx');
    });