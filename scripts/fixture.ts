import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import ExcelJS from "exceljs";

/**
 * Builds a fixture that mimics the client's messy monthly sheet: per-supplier
 * horizontal blocks, Data/Valor pairs, payment method buried in the item label,
 * an amounts-as-text block, and one already-normalized tab.
 */
export async function buildFixture(fixturePath: string): Promise<string> {
  mkdirSync(dirname(fixturePath), { recursive: true });
  const workbook = new ExcelJS.Workbook();

  const june = workbook.addWorksheet("junho26");

  june.getCell("A1").value = "MERCADO";
  june.getCell("A2").value = "Data";
  june.getCell("B2").value = "Valor";
  june.getCell("A3").value = new Date(Date.UTC(2026, 5, 2));
  june.getCell("B3").value = 1234.56;
  june.getCell("A4").value = new Date(Date.UTC(2026, 5, 9));
  june.getCell("B4").value = 890.1;

  june.getCell("D1").value = "BEBIDAS";
  june.getCell("D2").value = "Item";
  june.getCell("E2").value = "Valor";
  june.getCell("D3").value = "ambev (pix)";
  june.getCell("E3").value = 320.5;
  june.getCell("D4").value = "coca lata (dinheiro)";
  june.getCell("E4").value = 88.9;

  june.getCell("G1").value = "QUITANDA";
  june.getCell("G2").value = "Item";
  june.getCell("H2").value = "Valor";
  june.getCell("G3").value = "ovo caipira";
  june.getCell("H3").value = "R$ 1.234,56";
  june.getCell("G4").value = "cheiro verde";
  june.getCell("H4").value = "45,90";

  june.getCell("J1").value = "FATURAMENTO";
  june.getCell("J2").value = "Data";
  june.getCell("K2").value = "Valor";
  june.getCell("J3").value = "02/06";
  june.getCell("K3").value = 5200;
  june.getCell("J4").value = "10/06";
  june.getCell("K4").value = 4300;

  const target = workbook.addWorksheet("Lançamentos");
  target.addRow(["Data", "Tipo", "Categoria", "Descrição", "Valor", "Conta", "Status"]);
  target.addRow(["2026-06-02", "entrada", "Faturamento", "Vendas do dia", 5200, "Dinheiro", "pago"]);
  target.addRow(["2026-06-03", "saida", "Insumos", "Mercado Central", 1234.56, "PIX", "pago"]);
  target.addRow(["2026-06-05", "saida", "Bebidas", "Ambev", 320.5, "PIX", "pendente"]);
  target.addRow(["2026-06-06", "saida", "Gás", "Botijão", 130, "Dinheiro", "pago"]);

  await workbook.xlsx.writeFile(fixturePath);
  return fixturePath;
}
