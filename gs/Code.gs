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
	
    // Discussion of DTASOF tomorrow: 
    // https://pocketsense.blogspot.com/2010/08/replacing-microsoft-money-continued.html	
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