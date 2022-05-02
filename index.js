const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");
const data = require("./data.json");

const formatData = (data = {}) => {
  const tam = data.listaDeProdutos.length;
  const num1 = Math.ceil(tam / 2);

  const listaNova = [];
  let pos = 0;

  data.listaDeProdutos.forEach((prod, index) => {
    index === num1 && (pos = 0);
    if (index < num1) {
      listaNova[pos] = {
        ...listaNova[pos],
        produto1: `${prod.quantidade} ${prod.nome} (${
          prod.valorUnitario
        }) - ${prod.valorTotal.toFixed(2)}`,
      };
    } else {
      listaNova[pos] = {
        ...listaNova[pos],
        produto2: `${prod.quantidade} ${prod.nome} (${
          prod.valorUnitario
        }) - ${prod.valorTotal.toFixed(2)}`,
      };
    }
    pos++;
  });

  return listaNova;
};

const content = fs.readFileSync(
  path.resolve("./docs", "contrato.docx"),
  "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

doc.render({
  nome: data.nomeCliente,
  phone: "44 99700-0617",
  cpf: "117.645.989-97",
  dataDeEntrega: data.dataEntrega,
  local: data.enderecoDeEntrega,
  logradouro: "Rua mario monteschio, 436",
  total: data.valorTotal,
  quantidade: data.quantidadeTotal,
  produtos: formatData(data),
});

const buf = doc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});

fs.writeFileSync(path.resolve("./docs", "output.docx"), buf);
