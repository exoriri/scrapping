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


        for (let i = 8; i < websites.length; i++) {
            const contacts = {
                tel: '#N/A',
                vk: '#N/A'
            };

            if (phoneNumbers[i].error) {
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
                    console.log('trying http');
                    try {
                        const tel = await getScrappedData(websites[i], 'http');

                        if (tel.length === 0) {
                            const tel = await getScrappedData(`${websites[i]}/contacts`, 'http');
                            if (tel.length === 0) {
                                const tel = await getScrappedData(`${websites[i]}/contact`, 'http');
                                contacts.tel = tel;
                            } else {
                                contacts.tel = tel;
                            };
                        } else {
                            contacts.tel = tel;
                        };
                    } catch(e) {
                        console.log('TRIED ALL VARIANTS')
                    }
                }
            } else {
                contacts.tel = phoneNumbers[i];
            }

            res[websites[i]] = contacts;

            console.log('RESULT', res);

        };

        for (let i = 2; i < Object.keys(res).length; i++) {
            contactsWorksheet.addRow({
                rank: ranks[i],
                domain: websites[i],
                organicKeywords: organicKeywords[i],
                organicTraffic: organicTraffics[i],
                organicCost: organicCost[i],
                adwordsKeywords: adwordsKeywords[i],
                adwordsTraffic: adwordsTraffics[i],
                adwordsCost: adwordsCosts[i],
                phoneNumber: res[Object.keys(res)[i]].tel.length !== 0 ? res[Object.keys(res)[i]].tel.length !== 0 : '#N/A',
                anotherContact: additionalInfo[i],
                name: names[i],
                comment: '',
                status: statuses[i]
            })
        }



        const contactsXlsx = await workbookWriting.xlsx.writeFile('contacts.xlsx');
    });