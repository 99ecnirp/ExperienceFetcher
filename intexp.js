//AIM :- To create an excel file containing the name and link of the interview experiences of various companies.

let fs = require("fs");
let path = require("path");
let xlsx = require("xlsx");
const puppeteer = require("puppeteer");
let page, browser;

let dirPath = path.join(__dirname,"InterviewExperience");
isDirrectory(dirPath);


(async function fn() {
    let browserStartPromise = await puppeteer.launch({
        headless: false, defaultViewport: null,
        args: ["--start-maximized","--disable-notifications"],
    })

    let browserObj = await browserStartPromise;

    console.log("--------------------------------");
    console.log("First Tab Opened");

    browser = browserObj
    let  pages = await browser.pages();
    page = pages[0];
    await page.setDefaultNavigationTimeout(0);

    await page.goto("https://www.google.com");
    await page.type("input[title='Search']","www.geeksforgeeks.org",{delay : 200});
    await page.keyboard.press("Enter",{delay : 100});
    await page.waitForSelector(".LC20lb.DKV0Md",{visible : true});
    await page.click(".LC20lb.DKV0Md");      

    
    await page.waitForTimeout(4000);
    await page.waitForSelector("span.close",{visible:true});

    await page.click("span.close",{visible:true});
    console.log("--------------------------------");
    console.log("Pop up finished");

    await page.waitForSelector(".gfg-icon.gfg-icon_arrow-down.gfg-icon_header");
    let arr1 =  await page.$$(".gfg-icon.gfg-icon_arrow-down.gfg-icon_header");
    let value1 = await page.evaluate(function(element){
        element.click();
    },arr1[0]);
    console.log("--------------------------------");
    console.log("Dropdown tutorial arrow clicked");
    await page.waitForTimeout(3000);
   
    await page.waitForSelector(".mega-dropdown__list-item .gfg-icon.gfg-icon_arrow-right");
    let arr2 =  await page.$$(".mega-dropdown__list-item .gfg-icon.gfg-icon_arrow-right");
    let value2 = await page.evaluate(function(element){
        element.click();
    },arr2[4]);
    console.log("--------------------------------");
    console.log("Dropdown Interview Corner clicked");
    await page.waitForTimeout(3000);
     
    await page.waitForSelector(".mega-dropdown .mega-dropdown__list-item a");
    let arr3 =  await page.$$(".mega-dropdown .mega-dropdown__list-item a");
    let value3 = await page.evaluate(function(element){
        element.click();
    },arr3[44]); 
    console.log("--------------------------------");
    console.log("Interview Experience clicked");


    await page.waitForSelector(".sUlClass .sLiClass a");
    let nameOfCompanyArr = await page.$$(".sUlClass .sLiClass a") ;
    console.log("--------------------------------");
    console.log("Total Companies in the page :- ",nameOfCompanyArr.length);    
        
    let companyLinksArr=[] ;
    for (let i = 0; i < nameOfCompanyArr.length; i++) {
        let linkPromise = await page.evaluate(function (elem) {
            return elem.getAttribute("href");
        }, nameOfCompanyArr[i]);
        //console.log(linkPromise);
        companyLinksArr.push(linkPromise);
    }
    await page.waitForTimeout(4000);
    await autoScroll(page);
    await page.waitForTimeout(4000);

    //You can run it for all the 411 comapnies if you want, all edge cases are handled 
    for(let i = 0; i < companyLinksArr.length ; i++){
        await getExperience(companyLinksArr[i],i+1);
    }
        
    page.close();    
    browser.close();
    console.log("--------------------------------");
    console.log("Task has been successfully finished.")
    console.log("--------------------------------");
    
})();

async function getExperience( cmpLink, count){
    page = await browser.newPage();
    // await page.setDefaultNavigationTimeout(0);
    
    await page.goto(cmpLink,{visible : true });
    await page.waitForTimeout(3000);
    
    const exists = await page.$eval(".archive-title span",() => true).catch(() => false);
    //Have do do this as there is a tag where no data is present
    if(exists == true){
        await page.waitForSelector(".archive-title span",{visible:true});

        let elementName = await page.$(".archive-title span",{visible:true});
        let companyName = await page.evaluate(getTextElem,elementName);
        //console.log(typeof companyName);

        let star = companyName.includes("*");
        // console.log("value of star : " + star);
        //I have done this to handle a particular case where the name of comapny is 24*7; the problem here is that one can not create a file with
        // " * " in its name; i ahve exchanged " * " with " X "(see line 128);
        
        console.log("--------------------------------");
        console.log("(" + count + ") " + "Name of the current company :- ",companyName);
        console.log("--------------------------------");

    
        await page.waitForSelector(".content .head a");
        let expArr = await page.$$(".content .head a") ;
        
        if(star){
            let originalName = companyName.split("*");
            let firstPart = originalName[0];
            let secondPart = originalName[1];
            companyName = `${firstPart}X${secondPart}`;
        }
        

        let folderPath = path.join(__dirname,"InterviewExperience",companyName);
        isDirrectory(folderPath);

        let filePath = path.join(folderPath,companyName+".xlsx");
        let content  = excelReader(filePath,companyName); 
        let array = [];

        for(let i = 0;i < expArr.length; i++){
            let expLink = await page.evaluate(function (elem) {
                return elem.getAttribute("href");
            }, expArr[i]);
            let heading = await page.evaluate(getTextElem,expArr[i]);
            console.log( heading + "  ------->  " + expLink);

            let interViewExp = {
                name : heading,
                link : expLink
            }
            array.push(interViewExp);
            content.push(interViewExp);
        }

        console.table(array);
        excelWriter(filePath, content, companyName);
        page.close();
        console.log("--------------------------------------------------------------------------------------------------");
    }else{
       
        console.log("--------------------------------------------------------------------------------------------------");
        page.close();
    }    
};

function getTextElem(element){
    return element.textContent.trim();
}

async function autoScroll(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;

                if(totalHeight >= scrollHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
}

// async function scrollToBottom() {
//     await page.evaluate(goToBottom);
//     function goToBottom() {
//         window.scrollBy(0, window.innerHeight);
//         // console.log(window.innerHeight);
//          console.log("scrolled");
//     }
// }

function excelReader(filePath, name) {
    if (!fs.existsSync(filePath)) {
        return [];
    } else {
        let wb = xlsx.readFile(filePath);
        let excelData = wb.Sheets[name];
        let ans = xlsx.utils.sheet_to_json(excelData);
        return ans;
    }
}

function excelWriter(filePath, json, name) {
    
    let newWB = xlsx.utils.book_new();   
    let newWS = xlsx.utils.json_to_sheet(json);
   
    xlsx.utils.book_append_sheet(newWB, newWS, name);   
    xlsx.writeFile(newWB, filePath);
}
 
function isDirrectory(folderPath) {
    if(fs.existsSync(folderPath)== false){
        fs.mkdirSync(folderPath);
    }
}
