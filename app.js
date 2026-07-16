/**
 * app.js — East Coast Placard Generator
 *
 * Workflow:
 *  1. Populate the product <select> from the PRODUCTS catalog.
 *  2. On form submit, validate inputs, calculate pallet count,
 *     and render one .placard card per pallet inside #printArea.
 *  3. The Print button triggers window.print() so only placards
 *     are shown (form is hidden via @media print in styles.css).
 */

// ── Product catalog ───────────────────────────────────────────────
const PRODUCTS = [
  { name: "SHELLED EDAMAME 9OZ",          casesPerPallet: 160 },
  { name: "STUFFED PEPPERS WITH BEEF 15OZ", casesPerPallet: 143 },
  { name: "KUNG PAO BRUSSEL SPROUTS 14OZ",  casesPerPallet: 143 },
  { name: "GARLIC ASIAGO DIP 10 OZ",        casesPerPallet: 143 },
  { name: "CHICKEN STUFFED POBLANO 15 OZ",  casesPerPallet: 143 },
  { name: "ROASTED GREEN VEGETABLES 12 OZ", casesPerPallet: 143 }
];

// ── DOM references ────────────────────────────────────────────────
const form         = document.getElementById("placardForm");
const productSelect = document.getElementById("productSelect");
const printArea    = document.getElementById("printArea");
const btnPrint     = document.getElementById("btnPrint");

// ── Populate product <select> ─────────────────────────────────────
PRODUCTS.forEach(function (product, index) {
  const option = document.createElement("option");
  option.value = index;           // store array index as value
  option.textContent = product.name;
  productSelect.appendChild(option);
});

// ── Validation helpers ────────────────────────────────────────────

/**
 * Mark a field as valid or invalid and return whether it passed.
 * @param {string} fieldId  - id of the <input> or <select>
 * @param {boolean} isValid
 * @returns {boolean}
 */
function setValidity(fieldId, isValid) {
  const el = document.getElementById(fieldId);
  if (isValid) {
    el.classList.remove("invalid");
  } else {
    el.classList.add("invalid");
  }
  return isValid;
}

/**
 * Validate all form fields.
 * @returns {boolean} true if every field is valid.
 */
function validateForm() {
  const poNumber    = document.getElementById("poNumber").value.trim();
  const destination = document.getElementById("destination").value.trim();
  const shippingDate = document.getElementById("shippingDate").value;
  const productCode = document.getElementById("productCode").value.trim();
  const totalCasesRaw = document.getElementById("totalCases").value;
  const totalCases  = Number(totalCasesRaw);

  const v1 = setValidity("poNumber",    poNumber !== "");
  const v2 = setValidity("destination", destination !== "");
  const v3 = setValidity("shippingDate", shippingDate !== "");
  const v4 = setValidity("productCode", productCode !== "");
  // totalCases must be a positive whole number
  const v5 = setValidity(
    "totalCases",
    totalCasesRaw !== "" && Number.isInteger(totalCases) && totalCases >= 1
  );

  return v1 && v2 && v3 && v4 && v5;
}

// ── Placard builder ───────────────────────────────────────────────

/**
 * Build one placard <article> element.
 * @param {{
 *   palletNumber: number,
 *   totalPallets: number,
 *   poNumber: string,
 *   destination: string,
 *   shippingDate: string,
 *   productCode: string,
 *   product: {name: string, casesPerPallet: number}
 * }} data
 * @returns {HTMLElement}
 */
function buildPlacard(data) {
  const article = document.createElement("article");
  article.className = "placard";

  // Format shipping date for display (YYYY-MM-DD → MM/DD/YYYY)
  const dateParts = data.shippingDate.split("-");
  const displayDate = dateParts.length === 3
    ? dateParts[1] + "/" + dateParts[2] + "/" + dateParts[0]
    : data.shippingDate;

  article.innerHTML =
    '<h2 class="placard-title">TRADER JOE\'S</h2>' +

    '<div class="placard-row">' +
      '<strong>DESTINATION:</strong>' +
      '<span>' + escapeHtml(data.destination) + '</span>' +
    '</div>' +

    '<div class="placard-row">' +
      '<strong>TRADER JOE\'S PO#:</strong>' +
      '<span>' + escapeHtml(data.poNumber) + '</span>' +
    '</div>' +

    '<div class="placard-row">' +
      '<strong>SHIPPING DATE:</strong>' +
      '<span>' + escapeHtml(displayDate) + '</span>' +
    '</div>' +

    '<div class="placard-row">' +
      '<strong>CASES PER PALLET:</strong>' +
      '<span>' + data.product.casesPerPallet + ' CS</span>' +
    '</div>' +

    '<div class="placard-row">' +
      '<strong>PALLET:</strong>' +
      '<span>(' + data.palletNumber + ') OF (' + data.totalPallets + ')</span>' +
    '</div>' +

    '<hr class="placard-divider" />' +

    '<p class="placard-product-name">' + escapeHtml(data.product.name) + '</p>' +

    '<div class="placard-row">' +
      '<strong>CODE:</strong>' +
      '<span>' + escapeHtml(data.productCode) + '</span>' +
    '</div>' +

    '<div class="placard-unique-id">' +
      '<strong>Unique ID:</strong>' +
      '<span class="unique-id-blank" aria-label="Unique ID — fill manually"></span>' +
    '</div>';

  return article;
}

/**
 * Minimal HTML-escape to prevent XSS from form inputs.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#x27;");
}

// ── Form submit ───────────────────────────────────────────────────
form.addEventListener("submit", function (event) {
  event.preventDefault();   // stop browser's native form submission

  if (!validateForm()) {
    return;                  // stop if any field is invalid
  }

  const poNumber     = document.getElementById("poNumber").value.trim();
  const destination  = document.getElementById("destination").value.trim();
  const shippingDate = document.getElementById("shippingDate").value;
  const productCode  = document.getElementById("productCode").value.trim();
  const totalCases   = Number(document.getElementById("totalCases").value);
  const product      = PRODUCTS[Number(productSelect.value)];

  // Calculate total pallets (ceiling division)
  const totalPallets = Math.ceil(totalCases / product.casesPerPallet);

  // Clear previous placards
  printArea.innerHTML = "";

  // Generate one placard per pallet
  for (let palletNumber = 1; palletNumber <= totalPallets; palletNumber++) {
    const placard = buildPlacard({
      palletNumber,
      totalPallets,
      poNumber,
      destination,
      shippingDate,
      productCode,
      product
    });
    printArea.appendChild(placard);
  }

  // Enable the Print button now that placards exist
  btnPrint.disabled = false;
  // Scroll to the first placard
  printArea.scrollIntoView({ behavior: "smooth", block: "start" });
});

// ── Print button ──────────────────────────────────────────────────
btnPrint.addEventListener("click", function () {
  window.print();
});
