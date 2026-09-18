import { stripAccents } from "./parse.js";
import type { Cell } from "./types.js";

const PAYMENT_PATTERNS: Array<{ label: string; re: RegExp }> = [
  { label: "Dinheiro", re: /dinheiro|especie|cash/i },
  {
    label: "Cartão",
    re: /cart[aã]o|cr[eé]dito|d[eé]bito|maquininha|maquina|cielo|stone|pagseguro|infinitepay|getnet/i,
  },
  { label: "Boleto", re: /boleto|ted\b|doc\b|transfer[eê]ncia|autom[aá]tico|internet banking/i },
  { label: "PIX", re: /\bpix\b/i },
];

export interface PaymentHint {
  conta: string;
  /** The original label with the payment annotation removed. */
  cleaned: string;
  /** What matched, for auditing. */
  matched: string;
}

/**
 * The legacy Doce Sabor sheet encodes the payment method inside the item label,
 * e.g. "ovo caipira (dinheiro)" or "ambev - pix". This pulls it out so the new
 * schema's `conta` column can be filled and the description normalized.
 *
 * Order matters: "cartao de credito" must win over a bare "credito", and
 * "debito automatico" is a boleto, not a card — hence the ordered patterns.
 */
export function extractPaymentHint(raw: Cell | undefined): PaymentHint | null {
  if (typeof raw !== "string") return null;
  const text = raw;
  const normalized = stripAccents(text.toLowerCase());
  const findHit = (candidate: string) =>
    PAYMENT_PATTERNS.find((p) => p.re.test(stripAccents(candidate.toLowerCase())));

  // 1) Parenthesised annotation anywhere in the label.
  const paren = /[([{]([^)\]}]{2,40})[)\]}]/g;
  let match: RegExpExecArray | null;
  while ((match = paren.exec(text)) !== null) {
    const hit = findHit(match[1] ?? "");
    if (hit) {
      const start = match.index;
      const cleaned = (text.slice(0, start) + text.slice(start + match[0].length))
        .replace(/\s{2,}/g, " ")
        .replace(/[\s\-–—,;/]+$/, "")
        .trim();
      return { conta: hit.label, cleaned: cleaned || text, matched: match[0] };
    }
  }

  // 2) Trailing " - dinheiro" style annotation.
  const tail = /[\s\-–—/,;]+([^\-–—/,;]{2,40})$/.exec(text);
  if (tail) {
    const hit = findHit(tail[1] ?? "");
    if (hit) {
      const cleaned = text.slice(0, tail.index).replace(/[\s\-–—,;/]+$/, "").trim();
      if (cleaned) return { conta: hit.label, cleaned, matched: (tail[0] ?? "").trim() };
    }
  }

  // 3) Bare mention anywhere in the label (weakest signal).
  if (normalized.length > 0) {
    const hit = PAYMENT_PATTERNS.find((p) => p.re.test(normalized));
    if (hit) {
      const cleaned = text
        .replace(/[([{][^)\]}]*[)\]}]/g, "")
        .replace(/\s{2,}/g, " ")
        .trim();
      return { conta: hit.label, cleaned: cleaned || text, matched: hit.label };
    }
  }

  return null;
}
