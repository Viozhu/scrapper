function fetchData(e) {

  var listado = {
    "Trezor One": [
      "BF2",
      "BF24"
    ],
    "Trezor T": [
      "BF2",
      "BF15"
    ],
    "Ledger Nano S": [
      "BG2",
      "BG38"
    ],
    "Ledger Nano S Plus": [
      "BF2",
      "BF15"
    ],
    "Ledger Nano X": [
      "BF2",
      "BF19"
    ],
    "Safepal S1": [
      "BF2",
      "BF14"
    ]
  }
  const sh = SpreadsheetApp.openById("1sYzaQBPfyU3RulnS-JPEy3WWJojs2bg8G_h8YgOuU44")

  for (const property in listado) {


    var sheet = sh.getSheetByName([`${property}`]);
    var range = sheet.getRange(String(`${listado[property][0]}`) + ':' + String(`${listado[property][1]}`));
    var valuesUrl = range.getValues().map(elemento => elemento[0]).filter(elm => elm !== '')

    var options = {
      "method": "GET",
      "followRedirects": false,
      "muteHttpExceptions": true,
      "responseType": "text"
    };

    for (var indexAll = 0; indexAll < valuesUrl.length; indexAll++) {
      var data = UrlFetchApp.fetch(valuesUrl[indexAll], options).getContentText().replace(/(\r\n|\n|\r|\t)/gm, " ").split('>')
      const obj = {
        name: '',
        price: '',
        status: true,
        seller: '',
        sales: '',
        invoice: '',
        warranty: '',
        fee: '',
        url: '',
      }

      data.forEach(function (part, index) {

        var text = part.split('<')[0]
        if (text.replace(/([ ]+)/gm, "") != '') {

          // if(part.includes('FACTURA') || part.includes('factura') || part.includes('Factura') ){
          //   console.log(part.split)
          // }

          if (part.includes('Nuevo  | ')) {
            // Nuevo  |  11002 vendidos</span
            const sales = part.replace('Nuevo  |', '').replace('</span', '')
            obj.sales = sales.replace('vendidos','').replace('vendido','')
          }

          if (part.includes('meses de garantía')) {
            obj.warranty = part.replace('.</p', '')
          }

          if (part.includes('cuotas sin interés</span')) {
            obj.fee = part.replace('cuotas sin interés</span', 'cuotas s/interés').replace('hasta', '')
          }

          if (part.includes(`{"seller_id"`)) {
            const cutString = part.split('melidata("add", "event_data",')[1]
            const cutStingBetter = cutString?.split('); melidata')[0]
            try {
              const userData = JSON.parse(cutStingBetter)
              obj.seller = userData.seller_id
            } catch (error) {
              console.log(error)

            }

          }

          if (part.includes('Publicación pausada')) {
            obj.status = false
          }
          if (part.includes(`"offers":{"price"`)) {
            try {
              const newData = JSON.parse(part.split('<')[0].replace(/([ ]{2,})/gm, " ").replace(/(^ )/gm, "").replace(`(function(win, doc){     function loadScripts (s) { if (!s || !s.length) return; s = s.slice(0); var h = doc.head || doc.getElementsByTagName('head')[0]; var cbStack = {}; var cbChild = {}; for (var i = 0; i `, ''))
              
              obj.name = newData.name
              obj.price = newData.offers.price
              obj.url = newData.offers.url
            } catch (error) {
              console.log(error)
            }


          }
        }
        if (data.length === (index + 29)) {

          const day = new Date()
          const getFullDate = `${day.getDate()}/${day.getMonth()}/${day.getFullYear()}`

          // sh.appendRow([ obj.name, `$${obj.price}`, obj.status ? 'Activa' : 'Pausada', obj.seller, obj.warranty, obj.fee,obj.sales, obj.url])
          var range = sheet.getDataRange();
          var values = range.getValues();

          var header = []
          values[0].forEach(e => {
            if (e == '') return null
            header.push(e)
          })
          const abc = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ']

          const lastRowValue = sheet.getRange(`${abc[header.length]}1`)

          const getDate = `- ${day.getDate()}/${day.getMonth()} -`

         
          if (header[header.length - 1] == getDate) {
            for (var i = 0; i < valuesUrl.length; i++) {
              const currentRow = sheet.getRange(`${abc[header.length - 1]}${indexAll + 2}`)
              const setPrice = sheet.getRange(`B${indexAll + 2}`)
              const setStatus = sheet.getRange(`C${indexAll + 2}`)

              currentRow.setValue(obj.sales ? obj.sales : '0')
              setPrice.setValue(obj.price)
              setStatus.setValue(obj.status ? 'Activo' : 'Pausado')
            }
          } else {
            lastRowValue.setValue(getDate)
            for (var i = 0; i < valuesUrl.length; i++) {
              const currentRow = sheet.getRange(`${abc[header.length]}${indexAll + 2}`)
              const setPrice = sheet.getRange(`B${indexAll + 2}`)
              const setStatus = sheet.getRange(`C${indexAll + 2}`)

              currentRow.setValue(obj.sales ? obj.sales : '0')
              setPrice.setValue(obj.price)
              setStatus.setValue(obj.status ? 'Activo' : 'Pausado')

            }
          }
        }

      }
      )

    }

  };
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

 //Full object
// sh.appendRow([part.split('<')[0].replace(/([ ]{2,})/gm," ").replace(/(^ )/gm,"").replace(`(function(win, doc){ function loadScripts (s) { if (!s || !s.length) return; s = s.slice(0); var h = doc.head || doc.getElementsByTagName('head')[0]; var cbStack = {}; var cbChild = {}; for (var i = 0; i `,'')])
