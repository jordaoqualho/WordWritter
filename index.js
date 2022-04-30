const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

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
  nome: "Italo Qualho da Silva",
  phone: "44 99700-0617",
  cpf: "117.645.989-97",
  dataDeEntrega: "30/04/2022",
  local: "Giardino",
  logradouro: "Rua mario monteschio, 436",
  total: 3500,
  produtos: [
    { nome: "bombom de morango", valor: 3.2 },
    { nome: "bombom de uva", valor: 3.2 },
    { nome: "bombom de coxinha", valor: 3.2 },
    { nome: "bombom de amendoas", valor: 3.2 },
    { nome: "bombom de morango", valor: 3.2 },
  ],
});

doc.render();

const buf = doc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});

fs.writeFileSync(path.resolve("./docs", "output.docx"), buf);
