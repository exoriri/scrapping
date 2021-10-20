const { default: axios } = require('axios');
const cheerio = require('cheerio');
const Excel = require('exceljs');
const fs = require('fs');
const { domain } = require('process');
const puppeteer = require('puppeteer');

const workbook = new Excel.Workbook();
const workbookWriting = new Excel.Workbook();

const getScrappedData = async (website) => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    try {
        await page.goto(`https://${website}/`);

        let phones = await page.evaluate(() => {
            let res = '';

            const links = document.querySelectorAll('a[href^="tel:"]');

            links.forEach(link => {
                res += `${link.getAttribute('href').replace(/tel:/, '')}, `;
            });

            return res;    
        });

        if (!phones) {
            await (await page.$('a[href="/contacts"]')).click();
            await page.waitForNavigation();  
            let contactsNums = await page.evaluate(() => {
                let res = '';
    
                const links = document.querySelectorAll('a[href^="tel:"]');
    
                links.forEach(link => {
                    res += `${link.getAttribute('href').replace(/tel:/, '')}, `;
                });
    
                return res;    
            });

            await browser.close();

            return contactsNums;
        } else {
            await browser.close();
            return phones;
        }

        
    } catch(e) {
        console.warn(e.message)
        return '';
    }

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
                {key:"rank", header:"Rank"}, 
                {key: "domain", header: "Domain"}, 
                {key: "organicKeywords", header: "Organic Keywords"},
                {key: "organicTraffic", header: "Organic Traffic"},
                {key: "organicCost", header: "Organic Cost"},
                {key: "adwordsKeywords", header: "Adwords Keywords"},
                {key: "adwordsTraffic", header: "Adwords Traffic"},
                {key: "adwordsCost", header: "Adwords Cost"},
                {key: "phoneNumber", header: "Phone number"},
                {key: "anotherContact", header: "Another contact"},
                {key: "name", header: "Name"},
                {key: "comment", header: "Comment"},
                {key: "status", header: "Status"}
            ];

            for (let i = 2; i < websites.length; i++) {
                const contacts = {
                    tel: '#N/A',
                    vk: '#N/A'
                };
                
                if (phoneNumbers[i].error) {
                    const number = await getScrappedData(websites[i]);
                    contacts.tel = number;
                };
                res[websites[i]] = contacts;   
                console.log(websites[i])
                console.log(contacts);
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
                    phoneNumber: res[Object.keys(res)[i]].tel, 
                    anotherContact: additionalInfo[i],
                    name: names[i],
                    comment: '',
                    status: statuses[i]
                })
            }



            const contactsXlsx = await workbookWriting.xlsx.writeFile('contacts.xlsx');
        });

 

                // try {
                //     const request = await axios(`https://${websites[i]}/`);
                //     const $ = cheerio.load(request.data);
                //     const tel = $('a[href^="tel:"]').attr("href");
                //     const vk = $('a[href^="https://vk."]').attr("href");

                //     if (tel) {
                //         contacts.tel = tel.replace('tel:', '');
                //     };

                //     if (vk) {
                //         contacts.vk = vk;
                //     };
                //     res[websites[i]] = contacts;

                //     console.log(res)
                // } catch(e) {
                //     try {
                //         const request = await axios(`http://${websites[i]}/`);
                //         const $ = cheerio.load(request.data);
                //         const tel = $('a[href^="tel:"]').attr("href");
                //         const vk = $('a[href^="https://vk."]').attr("href");
    
                //         if (tel) {
                //             contacts.tel = tel.replace('tel:', '');
                //         };
    
                //         if (vk) {
                //             contacts.vk = vk;
                //         };
                //         res[websites[i]] = contacts;
    
                //         console.log(res)
                //     } catch(e) {
                        
                //         contacts.tel = "Не получилось загрузить сайт";
                //         contacts.vk = "Не получилось загрузить сайт";
                //         res[websites[i]] = contacts;
                //     }
                // }
