<!-- Markdown reference: https://guides.github.com/features/mastering-markdown/ -->

# *Google_Sheet_Script_to_OFX*

This [Google Script App](https://www.google.com/script/start/) will transform security data in a Google Sheet into an [OFX formatted file](http://moneymvps.org/faq/article/8.aspx).  

The OFX file can then be imported into [Microsoft Money Plus Sunset](https://www.microsoft.com/en-us/download/details.aspx?id=20738) to update the portfolio's stock and mutual fund prices.

With this Google Script App you have a reliable, free source of stock and mutual fund data to keep your Microsoft Money portfolio up to date.

## Requirements

Implementing the Google Script App requires an authorization step.  See the article [***Authorization for Google Services***](https://developers.google.com/apps-script/guides/services/authorization) for details.  

For an alternative that does not require implementing the Google Script App, see [Microsoft Money Sunset Edition Open Financial Exchange (OFX) file for Updating Portfolio Security Prices](https://observablehq.com/@mariodelgadosr/microsoft-money-sunset-edition-open-file-exchange-ofx-file).

## Microsoft Money Stock Price Importing Background:

The articles listed here are for background and are **not** detailing software install prerequisites.  

The Google Script App is an alternative to the Python script detailed in the first article.  

* [Download Price Quotes to Microsoft Money After Microsoft Pulls the Plug](https://thefinancebuff.com/security-quote-script-for-microsoft-money.html)
* [Obtain stock and fund quotes after July 2013](http://moneymvps.org/faq/article/651.aspx)
* [Replacing Microsoft Money, Part 5: OFX Scripts](https://thefinancebuff.com/replacing-microsoft-money-part-5-ofx-scripts.html)


## Instructions

* Create a *Dummy* investment account, as detailed in: [Download Price Quotes to Microsoft Money After Microsoft Pulls the Plug](https://thefinancebuff.com/security-quote-script-for-microsoft-money.html). 
   
   This will be the Microsoft Money account used to import the security prices.  **Only the Microsoft Money portfolio stock prices will be updated**.  The number of securities holdings is set to zero by the VBA program as detailed in [this strategy](https://thefinancebuff.com/replacing-microsoft-money-part-5-ofx-scripts.html#comment-2748).
   
   **Note**: You don't need to do any of the Python setup in the [article](https://thefinancebuff.com/security-quote-script-for-microsoft-money.html), only the *Dummy* account setup. This VBA program is an alternative to the Python scripts and utilizes
   Morningstar Portfolio data as the security price source.
   
* Create a Google Sheet formatted similar to this [sample sheet](https://docs.google.com/spreadsheets/d/1e53kLtCKlcAKVyEFZOi0N7ABj0XWTScY1SGMqWEb4eA). 

	The sheet must contain the following information:

	Column | Content (***n***=row number)
	-------| -----------
	A | Security Ticker Symbol
	B | =GOOGLEFINANCE(A***n***)
	C | =If(IFNA(GOOGLEFINANCE(A***n***,"expenseratio"),false),"mutual fund","stock")
	
	![Screen Shot of Google Sheet](https://github.com/MarioDelgadoSr/Google_Sheet_Script_to_OFX/blob/master/img/sheet.png)
	

* [Add](https://zapier.com/learn/google-sheets/google-apps-script-tutorial/) [Code.gs](https://github.com/MarioDelgadoSr/Google_Sheet_Script_to_OFX/gs/Code.gs) Google Script App to the Google Sheet.

  Note: The script's ***onOpen()*** function will run every time the Google Sheet is opened.  If Google Sheet with the security data is already open, close it and re-open to add the ***Util*** menu option.

* As noted, the Google App Script will add a ***Util*** menu option with a ***Generate quotes.ofx file*** sub-menu option.  Selecting this option will create the file ***quotes.ofx** in your [Google Drive](https://www.google.com/drive/) root folder along with a button to download the file.

![Screen Shot of generate OFX file](https://github.com/MarioDelgadoSr/Google_Sheet_Script_to_OFX//blob/master/img/generate.png)

![Screen Shot of download OFX file](https://github.com/MarioDelgadoSr/Google_Sheet_Script_to_OFX/blob/master/img/download.png)

* When Microsoft Money was installed, it created a [file association](https://blogs.technet.microsoft.com/windowsinternals/2017/10/25/windows-10-how-to-configure-file-associations-for-it-pros/) for .ofx files with the [Microsoft Money Import Handler](http://moneymvps.org/faq/article/407.aspx).  
  
The [mimetype](https://developer.mozilla.org/en-US/docs/Web/API/Blob/type) associated with download has been assigned to [application/x-msmoney](https://slick.pl/kb/htaccess/complete-list-mime-types/). 
If your browser ([Chrome](https://www.techjunkie.com/automatically-open-downloads-chrome/), [Edge](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_win10-msoversion_other/always-open-files-of-this-type/68910598-cc0c-4603-81ea-acb1902109f1), [FireFox](https://support.mozilla.org/en-US/kb/change-firefox-behavior-when-open-file)) is properly configured for opening an OFX file, the ***Money*** application will automatically initiate and import the file with the [Microsoft Money Import Handler](http://moneymvps.org/faq/article/407.aspx).  
Otherwise, manually invoke the [Microsoft Money Import Handler](http://moneymvps.org/faq/article/407.aspx) to import the OFX file.  

* View your updated Microsoft Money Portfolio.


## Code.gs code:

````
// Documentation: https://github.com/MarioDelgadoSr/Google_Sheet_Script_to_OFX
// https://developers.google.com/apps-script/guides/html
// https://developers.google.com/apps-script/reference/drive/drive-app#createFile(String,String,String)
// https://developers.google.com/apps-script/reference/drive/file#setContent(String)
// https://developers.google.com/apps-script/reference/drive/file#getDownloadUrl()
// https://developers.google.com/apps-script/reference/html/html-service#createHtmlOutput(String)

function onOpen() {
  
  SpreadsheetApp.getUi() 
      .createMenu('Util')
      .addItem('Generate quotes.ofx file', 'downloadOFX')
      .addToUi();
}

function downloadOFX(){


  const ofx = ofxFile();

  const  chkFile = DriveApp.getFilesByName("quotes.ofx");
  let file;

  if (chkFile.hasNext()) {

    file = chkFile.next();
    file.setContent(ofx);

  }
  else {

    file = DriveApp.createFile("quotes.ofx", ofx, "application/x-msmoney");

  }


  const anchor = `<a href="${file.getDownloadUrl()}"><button>Download quotes.ofx file</button></a>`;

  SpreadsheetApp.getUi() 
      .showModalDialog(HtmlService.createHtmlOutput(anchor), 'Stock Quotes');


}

function ofxFile() {

    // Logic has to reference sheet with values pre-caluculated with GOOGLEFINANCE function
    // Sheet functions can't be referenced in JavaScript code: https://stackoverflow.com/a/54586150 
    // Open File Exchange specifications: https://www.ofx.net/downloads/OFX%202.2.pdf
    // Date format: page 89, 3.2.8.2  Date and Datetime YYYYMMDDHHMMSS GMT

    const strDate = new Date().toISOString().substring(0,19).replace(/T/g,"").replace(/-/g,"").replace(/:/g,"");
    const tomorrow = new Date();
          tomorrow.setDate(tomorrow.getDate() + 1);
    const strTomorrow = tomorrow.toISOString().substring(0,10).replace(/T/g,"").replace(/-/g,"");

    let ofx = startXML(strDate, strTomorrow);  
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];

    const range = sheet.getDataRange();
    const security = range.getValues();

    let ticker, price, secType;

    for (let i=0; i < security.length; i++){ 

        ticker =  security[i][0]; price = security[i][1]; secType = security[i][2];

        let posType = secType == "stock" ? posstock : posmf;

        ofx += posType(ticker, price, strDate); 


    }

    ofx += startInfo();

    for (let i=0; i < security.length; i++){ 

        ticker =  security[i][0]; price = security[i][1]; secType = security[i][2];

        let infoType = secType == "stock" ? stockinfo : mfinfo;

        ofx += infoType(ticker, price, strDate);

    }

    ofx += footer();
  
    ofx = header() + formatXml(ofx);
  
    return ofx;
} 




function header(){
  
const headerXML = 
`OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE`;
 
return headerXML;  

}  

function startXML(strDate, strTomorrow){
 
// https://developers.google.com/apps-script/reference/utilities/utilities#getUuid()

  const xml = 
`
<OFX>
<SIGNONMSGSRSV1>
<SONRS>
<STATUS>
<CODE>0</CODE>
<SEVERITY>INFO</SEVERITY>
<MESSAGE>Successful Sign On</MESSAGE>
</STATUS>
<DTSERVER>${strDate}</DTSERVER>                                                     
<LANGUAGE>ENG</LANGUAGE>
<DTPROFUP>20010918083000</DTPROFUP>
<FI>
<ORG>broker.com</ORG>
</FI>
</SONRS>
</SIGNONMSGSRSV1>
<INVSTMTMSGSRSV1>
<INVSTMTTRNRS>
<TRNUID>${Utilities.getUuid()}</TRNUID>
<STATUS>
<CODE>0</CODE>
<SEVERITY>INFO</SEVERITY>
</STATUS>
<CLTCOOKIE>4</CLTCOOKIE>
<INVSTMTRS>
<DTASOF>${strTomorrow}</DTASOF>
<CURDEF>USD</CURDEF>
<INVACCTFROM>
<BROKERID>dummybroker.com</BROKERID>
<ACCTID>0123456789</ACCTID>
</INVACCTFROM>
<INVPOSLIST>
`;
  
return xml;
  
}  
function posstock(strSecurity,price,strDate){
  
  const posstockXML =
`<POSSTOCK>
<INVPOS>
<SECID>
<UNIQUEID>${strSecurity}</UNIQUEID>
<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
</SECID>
<HELDINACCT>CASH</HELDINACCT>
<POSTYPE>LONG</POSTYPE>
<UNITS>0</UNITS>
<UNITPRICE>${price}</UNITPRICE>
<MKTVAL>${price}</MKTVAL>
<DTPRICEASOF>${strDate}</DTPRICEASOF>
</INVPOS>
</POSSTOCK>`;
  
  return posstockXML;  
  
}


function posmf(strSecurity,price,strDate){
  
  const posmfXML =   
`<POSMF>
<INVPOS>
<SECID>
<UNIQUEID>${strSecurity}</UNIQUEID>
<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
</SECID>
<HELDINACCT>CASH</HELDINACCT>
<POSTYPE>LONG</POSTYPE>
<UNITS>0</UNITS>
<UNITPRICE>${price}</UNITPRICE>
<MKTVAL>${price}</MKTVAL>
<DTPRICEASOF>${strDate}</DTPRICEASOF>
</INVPOS>
</POSMF>`;
  
  return posmfXML;  
}


function startInfo(){
 
  const startInfoXML =
`</INVPOSLIST>
</INVSTMTRS>
</INVSTMTTRNRS>
</INVSTMTMSGSRSV1>
<SECLISTMSGSRSV1>
<SECLIST>
`;
  
  return startInfoXML;
  
} 

function stockinfo(strSecurity, price){

const stockinfoXML = 
`<STOCKINFO>
<SECINFO>
<SECID>
<UNIQUEID>${strSecurity}</UNIQUEID>
<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
</SECID>
<SECNAME>${strSecurity}</SECNAME>
<TICKER>${strSecurity}</TICKER>
<UNITPRICE>${price}</UNITPRICE>
</SECINFO>
</STOCKINFO>`;

	return stockinfoXML
}

function mfinfo(strSecurity, price){

	const mfinfoXML = 
`<MFINFO>
<SECINFO>
<SECID>
<UNIQUEID>${strSecurity}</UNIQUEID>
<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
</SECID>
<SECNAME>${strSecurity}</SECNAME>
<TICKER>${strSecurity}</TICKER>
<UNITPRICE>${price}</UNITPRICE>
</SECINFO>
<MFTYPE>OPENEND</MFTYPE>
</MFINFO>`

    return mfinfoXML;
}	

function footer(){
 
 const footerXML = 
`</SECLIST>
</SECLISTMSGSRSV1>
</OFX>`;
  
  return footerXML;
  
} 

// https://stackoverflow.com/a/49458964
function formatXml(xml, tab) { // tab = optional indent value, default is tab (\t)
    var formatted = '', indent= '';
    tab = tab || '\t';
    xml.split(/>\s*</).forEach(function(node) {
        if (node.match( /^\/\w/ )) indent = indent.substring(tab.length); // decrease indent by one 'tab'
        formatted += indent + '<' + node + '>\r\n';
        if (node.match( /^<?\w[^>]*[^\/]$/ )) indent += tab;              // increase indent
    });
    return formatted.substring(1, formatted.length-3);
}
````


## Author

* **Mario Delgado**  Github: [MarioDelgadoSr](https://github.com/MarioDelgadoSr)
* LinkedIn: [Mario Delgado](https://www.linkedin.com/in/mario-delgado-5b6195155/)
* [My Data Visualizer](http://MyDataVisualizer.com): A data visualization application using the [*DataVisual*](https://github.com/MarioDelgadoSr/DataVisual) design pattern.


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details




