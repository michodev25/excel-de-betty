import seedCatalog from '../data/catalog.json';
import type { Catalog, Product } from '../types';

const STORAGE_KEY = 'listado-precios-catalog-v1';
const initialCatalog = seedCatalog as Catalog;

function required<T extends HTMLElement>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error(`Falta el elemento ${selector}`);
  return element;
}

const desktopRows = required<HTMLTableSectionElement>('#desktopRows');
const mobileList = required<HTMLDivElement>('#mobileList');
const searchInput = required<HTMLInputElement>('#searchInput');
const categoryFilter = required<HTMLSelectElement>('#categoryFilter');
const missingOnly = required<HTMLInputElement>('#missingOnly');
const dateInput = required<HTMLInputElement>('#listDate');
const importFile = required<HTMLInputElement>('#importFile');
const dialog = required<HTMLDialogElement>('#productDialog');
const form = required<HTMLFormElement>('#productForm');
const toast = required<HTMLDivElement>('#toast');
const deleteButton = required<HTMLButtonElement>('#deleteProduct');
const newCategoryField = required<HTMLElement>('#newCategoryField');
const originalNotes = required<HTMLElement>('#originalNotes');
const originalNotesContent = required<HTMLElement>('#originalNotesContent');

let catalog: Catalog = readCatalog();
let editingId: string | null = null;
let toastTimer: number | undefined;
let dateEdited = false;

function readCatalog(): Catalog {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const value = JSON.parse(raw) as Catalog;
      if (Array.isArray(value.categories) && Array.isArray(value.products)) return value;
    }
  } catch {
    // If browser storage is unavailable, the original catalog remains usable.
  }
  return structuredClone(initialCatalog);
}

function saveCatalog(): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(catalog));
  } catch {
    showToast('No se pudieron guardar los cambios en este navegador. Descarga una copia del Excel.');
  }
}

function showToast(message: string): void {
  toast.textContent = message;
  toast.classList.add('visible');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.remove('visible'), 3600);
}

function todayLocal(): string {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

function displayDate(iso: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : '—';
}

function updateDate(): void {
  const text = displayDate(dateInput.value);
  required<HTMLElement>('#datePreview').textContent = text;
  required<HTMLElement>('#mobileDatePreview').textContent = text;
}

function normalize(text: string): string {
  return text.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function parsePrice(raw: string): number | null | undefined {
  const value = raw.trim();
  if (!value) return null;
  if (!/^\d+(?:[.,]\d{1,2})?$/.test(value)) return undefined;
  const price = Number(value.replace(',', '.'));
  return Number.isFinite(price) && price <= 1_000_000_000 ? price : undefined;
}

function priceText(value: number | null): string {
  return value === null ? '' : value.toFixed(2).replace('.', ',');
}

function visibleProducts(): Product[] {
  const query = normalize(searchInput.value);
  const category = categoryFilter.value;
  const onlyMissing = missingOnly.checked;
  return catalog.products.filter((item) => {
    if (category && item.category !== category) return false;
    if (onlyMissing && item.price !== null) return false;
    if (!query) return true;
    return normalize(`${item.product} ${item.format} ${item.supplier} ${item.category}`).includes(query);
  });
}

function td(text: string, className = ''): HTMLTableCellElement {
  const cell = document.createElement('td');
  cell.textContent = text || '—';
  if (className) cell.className = className;
  return cell;
}

function priceInput(item: Product): HTMLInputElement {
  const input = document.createElement('input');
  input.className = 'price-input';
  input.type = 'text';
  input.inputMode = 'decimal';
  input.enterKeyHint = 'done';
  input.autocomplete = 'off';
  input.value = priceText(item.price);
  input.placeholder = 'Sin precio';
  input.dataset.priceId = item.id;
  input.setAttribute('aria-label', `Precio ($) de ${item.product}`);
  return input;
}

function priceError(): HTMLElement {
  const message = document.createElement('small');
  message.className = 'price-error';
  message.textContent = 'Usa hasta dos decimales o deja el campo vacío.';
  message.hidden = true;
  return message;
}

function editButton(item: Product, mobile = false): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = mobile ? 'card-edit' : 'edit-row';
  button.dataset.editId = item.id;
  button.textContent = mobile ? 'Editar datos' : 'Editar';
  button.setAttribute('aria-label', `Editar datos de ${item.product}`);
  return button;
}

function renderDesktop(items: Product[]): void {
  desktopRows.replaceChildren();
  const fragment = document.createDocumentFragment();
  for (const category of catalog.categories) {
    const group = items.filter((item) => item.category === category);
    if (!group.length) continue;
    const heading = document.createElement('tr');
    heading.className = 'category-row';
    heading.dataset.category = category;
    const cell = document.createElement('td');
    cell.colSpan = 7;
    cell.textContent = `${category} · ${group.length} ${group.length === 1 ? 'producto' : 'productos'}`;
    heading.append(cell);
    fragment.append(heading);
    for (const item of group) {
      const row = document.createElement('tr');
      row.className = 'item-row';
      row.dataset.productId = item.id;
      row.append(td(item.product, 'product-name'));
      row.append(td(item.format, !item.format ? 'cell-muted' : ''));
      row.append(td(item.unit, !item.unit ? 'cell-muted' : ''));
      const priceCell = document.createElement('td');
      priceCell.append(priceInput(item), priceError());
      row.append(priceCell);
      row.append(td(item.supplier, !item.supplier ? 'cell-muted' : ''));
      row.append(td(item.social, !item.social ? 'cell-muted' : ''));
      const actionCell = document.createElement('td');
      actionCell.append(editButton(item));
      row.append(actionCell);
      fragment.append(row);
    }
  }
  desktopRows.append(fragment);
}

function renderMobile(items: Product[]): void {
  mobileList.replaceChildren();
  const fragment = document.createDocumentFragment();
  for (const category of catalog.categories) {
    const group = items.filter((item) => item.category === category);
    if (!group.length) continue;
    const heading = document.createElement('h3');
    heading.className = 'mobile-category';
    heading.dataset.category = category;
    heading.textContent = category;
    fragment.append(heading);
    for (const item of group) {
      const card = document.createElement('article');
      card.className = 'product-card';
      card.dataset.productId = item.id;
      card.setAttribute('aria-labelledby', `product-${item.id}`);
      const top = document.createElement('div');
      top.className = 'card-top';
      const identity = document.createElement('div');
      const name = document.createElement('h4');
      name.className = 'card-name';
      name.id = `product-${item.id}`;
      name.textContent = item.product;
      const meta = document.createElement('div');
      meta.className = 'card-meta';
      meta.textContent = [item.format, item.unit].filter(Boolean).join(' · ') || 'Presentación por definir';
      identity.append(name, meta);
      top.append(identity);
      const supplier = document.createElement('div');
      supplier.className = 'card-supplier';
      supplier.textContent = item.supplier || 'Proveedor por definir';
      const bottom = document.createElement('div');
      bottom.className = 'card-bottom';
      const priceWrap = document.createElement('div');
      priceWrap.className = 'card-price-wrap';
      const label = document.createElement('label');
      label.textContent = 'Precio ($)';
      const input = priceInput(item);
      input.id = `mobile-price-${item.id}`;
      label.htmlFor = input.id;
      priceWrap.append(label, input, priceError());
      bottom.append(priceWrap, editButton(item, true));
      card.append(top, supplier, bottom);
      fragment.append(card);
    }
  }
  mobileList.append(fragment);
}

function renderCategoryChoices(): void {
  const selected = categoryFilter.value;
  categoryFilter.replaceChildren(new Option('Todas las categorías', ''));
  for (const category of catalog.categories) categoryFilter.add(new Option(category, category));
  categoryFilter.value = catalog.categories.includes(selected) ? selected : '';
}

function updateCounts(): void {
  const total = catalog.products.length;
  const priced = catalog.products.filter((item) => item.price !== null).length;
  required<HTMLElement>('#totalCount').textContent = String(total);
  required<HTMLElement>('#pricedCount').textContent = String(priced);
  required<HTMLElement>('#missingCount').textContent = String(total - priced);
  required<HTMLElement>('#categoryCount').textContent = String(catalog.categories.length);
}

function render(): void {
  updateCounts();
  const items = visibleProducts();
  required<HTMLElement>('#resultCount').textContent = `${items.length} ${items.length === 1 ? 'producto visible' : 'productos visibles'}`;
  required<HTMLElement>('#emptyState').hidden = items.length > 0;
  renderDesktop(items);
  renderMobile(items);
}

function formField<T extends HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(name: string): T {
  const field = form.elements.namedItem(name);
  if (!(field instanceof HTMLElement)) throw new Error(`Falta el campo ${name}`);
  return field as T;
}

function syncNewCategory(): void {
  const useNew = formField<HTMLSelectElement>('category').value === '__new__';
  newCategoryField.hidden = !useNew;
  formField<HTMLInputElement>('newCategory').required = useNew;
}

function openDialog(item?: Product): void {
  editingId = item?.id ?? null;
  form.reset();
  required<HTMLElement>('#dialogTitle').textContent = item ? 'Editar producto' : 'Añadir producto';
  required<HTMLElement>('#dialogEyebrow').textContent = item ? 'EDITAR DATOS' : 'NUEVO PRODUCTO';
  const categorySelect = formField<HTMLSelectElement>('category');
  categorySelect.replaceChildren(new Option('Selecciona una categoría', ''));
  for (const category of catalog.categories) categorySelect.add(new Option(category, category));
  categorySelect.add(new Option('+ Nueva categoría', '__new__'));
  formField<HTMLInputElement>('product').value = item?.product ?? '';
  categorySelect.value = item?.category ?? '';
  formField<HTMLInputElement>('format').value = item?.format ?? '';
  const unitSelect = formField<HTMLSelectElement>('unit');
  if (item?.unit && !Array.from(unitSelect.options).some((option) => option.value === item.unit)) unitSelect.add(new Option(item.unit, item.unit));
  unitSelect.value = item?.unit || 'UNIDAD';
  formField<HTMLInputElement>('price').value = item ? priceText(item.price) : '';
  formField<HTMLInputElement>('supplier').value = item?.supplier ?? '';
  formField<HTMLInputElement>('address').value = item?.address ?? '';
  const socialSelect = formField<HTMLSelectElement>('social');
  if (item?.social && !Array.from(socialSelect.options).some((option) => option.value === item.social)) socialSelect.add(new Option(item.social, item.social));
  socialSelect.value = item?.social ?? '';
  formField<HTMLTextAreaElement>('notes').value = item?.notes ?? '';
  deleteButton.hidden = !item;
  originalNotesContent.replaceChildren();
  const comments = Object.entries(item?.sourceComments ?? {});
  originalNotes.hidden = comments.length === 0;
  for (const [column, comment] of comments) {
    const paragraph = document.createElement('p');
    paragraph.textContent = `${column}: ${comment}`;
    originalNotesContent.append(paragraph);
  }
  syncNewCategory();
  dialog.showModal();
  requestAnimationFrame(() => formField<HTMLInputElement>('product').focus());
}

function onPriceChange(event: Event): void {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.dataset.priceId) return;
  const item = catalog.products.find((product) => product.id === input.dataset.priceId);
  if (!item) return;
  const parsed = parsePrice(input.value);
  if (parsed === undefined) {
    input.setAttribute('aria-invalid', 'true');
    const message = input.parentElement?.querySelector<HTMLElement>('.price-error');
    if (message) {
      message.id = `price-error-${item.id}-${input.closest('.mobile-list') ? 'mobile' : 'desktop'}`;
      message.hidden = false;
      input.setAttribute('aria-describedby', message.id);
    }
    showToast('Escribe un precio válido con hasta dos decimales, o deja el campo vacío.');
    return;
  }
  input.removeAttribute('aria-invalid');
  input.removeAttribute('aria-describedby');
  const message = input.parentElement?.querySelector<HTMLElement>('.price-error');
  if (message) message.hidden = true;
  item.price = parsed;
  saveCatalog();
  updateCounts();
  document.querySelectorAll<HTMLInputElement>('[data-price-id]').forEach((other) => {
    if (other.dataset.priceId === item.id && other !== input) other.value = priceText(parsed);
  });
  input.value = priceText(parsed);
  if (missingOnly.checked && parsed !== null) {
    for (const root of [desktopRows, mobileList]) {
      const entry = Array.from(root.querySelectorAll<HTMLElement>('[data-product-id]'))
        .find((element) => element.dataset.productId === item.id);
      entry?.remove();
      const group = visibleProducts().filter((product) => product.category === item.category);
      const heading = Array.from(root.querySelectorAll<HTMLElement>('[data-category]'))
        .find((element) => element.dataset.category === item.category);
      if (heading) {
        if (group.length === 0) heading.remove();
        else if (heading.matches('tr')) heading.firstElementChild!.textContent = `${item.category} · ${group.length} ${group.length === 1 ? 'producto' : 'productos'}`;
      }
    }
    const count = visibleProducts().length;
    required<HTMLElement>('#resultCount').textContent = `${count} ${count === 1 ? 'producto visible' : 'productos visibles'}`;
    required<HTMLElement>('#emptyState').hidden = count > 0;
  }
  showToast(parsed === null ? 'Precio eliminado.' : 'Precio actualizado.');
}

function onPriceInput(event: Event): void {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.dataset.priceId) return;
  if (parsePrice(input.value) !== undefined) {
    input.removeAttribute('aria-invalid');
    input.removeAttribute('aria-describedby');
    const message = input.parentElement?.querySelector<HTMLElement>('.price-error');
    if (message) message.hidden = true;
  }
}

function onEditClick(event: Event): void {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const button = target.closest<HTMLButtonElement>('[data-edit-id]');
  if (!button) return;
  const item = catalog.products.find((product) => product.id === button.dataset.editId);
  if (item) openDialog(item);
}

function createId(): string {
  return typeof crypto.randomUUID === 'function' ? crypto.randomUUID() : `new-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function saveProduct(event: SubmitEvent): void {
  event.preventDefault();
  if (!form.reportValidity()) return;
  const rawPrice = formField<HTMLInputElement>('price').value;
  const parsed = parsePrice(rawPrice);
  if (parsed === undefined) {
    formField<HTMLInputElement>('price').setCustomValidity('Usa un precio válido con hasta dos decimales.');
    form.reportValidity();
    formField<HTMLInputElement>('price').setCustomValidity('');
    return;
  }
  const selectedCategory = formField<HTMLSelectElement>('category').value;
  const category = selectedCategory === '__new__' ? formField<HTMLInputElement>('newCategory').value.trim() : selectedCategory;
  if (!category) return;
  const existing = editingId ? catalog.products.find((product) => product.id === editingId) : undefined;
  const item: Product = {
    id: existing?.id ?? createId(),
    category,
    product: formField<HTMLInputElement>('product').value.trim(),
    format: formField<HTMLInputElement>('format').value.trim(),
    unit: formField<HTMLSelectElement>('unit').value,
    price: parsed,
    supplier: formField<HTMLInputElement>('supplier').value.trim(),
    address: formField<HTMLInputElement>('address').value.trim(),
    social: formField<HTMLSelectElement>('social').value,
    notes: formField<HTMLTextAreaElement>('notes').value.trim(),
    ...(existing?.originalRow ? { originalRow: existing.originalRow } : {}),
    ...(existing?.sourceComments ? { sourceComments: existing.sourceComments } : {}),
  };
  if (!catalog.categories.includes(category)) catalog.categories.push(category);
  if (existing) Object.assign(existing, item);
  else catalog.products.push(item);
  saveCatalog();
  renderCategoryChoices();
  render();
  dialog.close();
  showToast(existing ? 'Producto actualizado.' : 'Producto añadido.');
}

async function onExport(): Promise<void> {
  if (!dateEdited) {
    dateInput.value = todayLocal();
    updateDate();
  }
  if (!dateInput.value || !/^\d{4}-\d{2}-\d{2}$/.test(dateInput.value)) {
    dateInput.focus();
    showToast('Elige una fecha válida para el listado.');
    return;
  }
  const invalid = Array.from(document.querySelectorAll<HTMLInputElement>('[data-price-id][aria-invalid="true"]'))
    .find((field) => field.getClientRects().length > 0);
  if (invalid) {
    invalid.scrollIntoView({ block: 'center' });
    invalid.focus();
    showToast('Corrige el precio marcado antes de descargar.');
    return;
  }
  const buttons = [required<HTMLButtonElement>('#exportTop'), required<HTMLButtonElement>('#exportMobile')];
  buttons.forEach((button) => { button.disabled = true; button.textContent = 'Preparando Excel…'; });
  try {
    const { exportWorkbook } = await import('../lib/workbook');
    await exportWorkbook(catalog.products, dateInput.value);
    showToast(`Excel del ${displayDate(dateInput.value)} descargado.`);
  } catch (error) {
    console.error('No se pudo exportar el Excel', error);
    showToast('No se pudo crear el Excel. Inténtalo de nuevo.');
  } finally {
    buttons[0].disabled = buttons[1].disabled = false;
    buttons[0].innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m0 0 4-4m-4 4-4-4M4 17v3h16v-3"/></svg> Descargar Excel';
    buttons[1].textContent = 'Descargar Excel';
  }
}

async function onImport(): Promise<void> {
  const file = importFile.files?.[0];
  if (!file) return;
  if (!window.confirm('El Excel importado sustituirá el catálogo guardado en este navegador. ¿Continuar?')) {
    importFile.value = '';
    return;
  }
  const button = required<HTMLButtonElement>('#importExcel');
  button.disabled = true;
  try {
    const { importWorkbook } = await import('../lib/workbook');
    catalog = await importWorkbook(file);
    saveCatalog();
    searchInput.value = '';
    missingOnly.checked = false;
    categoryFilter.value = '';
    renderCategoryChoices();
    render();
    showToast(`${catalog.products.length} productos importados.`);
  } catch (error) {
    console.error('No se pudo importar el Excel', error);
    showToast('No se pudo leer este Excel. Comprueba que tenga el formato del listado.');
  } finally {
    button.disabled = false;
    importFile.value = '';
  }
}

dateInput.value = todayLocal();
updateDate();
renderCategoryChoices();
render();

searchInput.addEventListener('input', render);
categoryFilter.addEventListener('change', render);
missingOnly.addEventListener('change', render);
dateInput.addEventListener('change', () => { dateEdited = true; updateDate(); });
desktopRows.addEventListener('change', onPriceChange);
mobileList.addEventListener('change', onPriceChange);
desktopRows.addEventListener('input', onPriceInput);
mobileList.addEventListener('input', onPriceInput);
desktopRows.addEventListener('click', onEditClick);
mobileList.addEventListener('click', onEditClick);
required<HTMLButtonElement>('#addProduct').addEventListener('click', () => openDialog());
required<HTMLButtonElement>('#closeDialog').addEventListener('click', () => dialog.close());
required<HTMLButtonElement>('#cancelDialog').addEventListener('click', () => dialog.close());
formField<HTMLSelectElement>('category').addEventListener('change', syncNewCategory);
form.addEventListener('submit', saveProduct);
deleteButton.addEventListener('click', () => {
  if (!editingId) return;
  const item = catalog.products.find((product) => product.id === editingId);
  if (!item || !window.confirm(`¿Eliminar «${item.product}» del listado?`)) return;
  catalog.products = catalog.products.filter((product) => product.id !== editingId);
  saveCatalog();
  render();
  dialog.close();
  showToast('Producto eliminado.');
});
required<HTMLButtonElement>('#exportTop').addEventListener('click', onExport);
required<HTMLButtonElement>('#exportMobile').addEventListener('click', onExport);
required<HTMLButtonElement>('#importExcel').addEventListener('click', () => importFile.click());
importFile.addEventListener('change', onImport);
