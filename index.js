const path = require("path");
const fs = require("fs");
var parser = require("xml2json");
var json2xls = require("json2xls");

getItemsInDirectory = async (dir) => {
  const directoryPath = path.join(__dirname, dir);
  const fileList = fs.readdirSync(directoryPath, function (err, files) {
    if (err) {
      return console.log("Unable to scan directory: " + err);
    }
    return files;
  });
  return Promise.resolve(fileList);
};

let counter = 1;
saveJsonResults = async (jsonContent) => {
  fs.writeFileSync(`./results/${counter}.json`, jsonContent, (err) => {
    return console.log(err);
  });
  counter = counter + 1;
};

const parseXmlToJson = (fileName) => {
  const result = fs.readFileSync(`./files/${fileName}`, "utf8");
  return parser.toJson(result);
};

const handleGenerateXLSX = (array) => {
  const productArray = [];

  array.map((file) => {
    const parsed = file;
    const fileProducts = parsed.nfeProc.NFe.infNFe.det;
    if (Array.isArray(fileProducts)) {
      fileProducts.map((prod) => {
        productArray.push({
          cEAN: prod.prod.cEAN,
          xProd: prod.prod.xProd,
          NCM: prod.prod.NCM,
          cEANTrib: prod.prod.cEANTrib,
          uCom: prod.prod.uCom,
          qCom: prod.prod.qCom,
          vUnCom: prod.prod.vUnCom,
          vProd: prod.prod.vProd,
          CNPJEmissor: parsed.nfeProc.NFe.infNFe.emit.CNPJ,
          NOMEEmissor: parsed.nfeProc.NFe.infNFe.emit.xNome,
          CNPJDestinatario: parsed.nfeProc.NFe.infNFe.dest.CNPJ,
          NOMEDestinatario: parsed.nfeProc.NFe.infNFe.dest.xNome,
        });
      });
    } else {
      productArray.push({
        cEAN: fileProducts.prod.cEAN,
        xProd: fileProducts.prod.xProd,
        NCM: fileProducts.prod.NCM,
        cEANTrib: fileProducts.prod.cEANTrib,
        uCom: fileProducts.prod.uCom,
        qCom: fileProducts.prod.qCom,
        vUnCom: fileProducts.prod.vUnCom,
        vProd: fileProducts.prod.vProd,
        CNPJEmissor: parsed.nfeProc.NFe.infNFe.emit.CNPJ,
        NOMEEmissor: parsed.nfeProc.NFe.infNFe.emit.xNome,
        CNPJDestinatario: parsed.nfeProc.NFe.infNFe.dest.CNPJ,
        NOMEDestinatario: parsed.nfeProc.NFe.infNFe.dest.xNome,
      });
    }
  });

  const xls = json2xls(productArray);
  fs.writeFileSync("final.xlsx", xls, "binary");
};

main = async () => {
  const files = await getItemsInDirectory("files");

  const jsonResultArray = files.map((file) => {
    const result = parseXmlToJson(file);
    return JSON.parse(result);
  });

  handleGenerateXLSX(jsonResultArray);

  return;
};

main();
