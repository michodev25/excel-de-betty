export interface Product {
  id: string;
  category: string;
  product: string;
  format: string;
  unit: string;
  price: number | null;
  supplier: string;
  address: string;
  social: string;
  notes: string;
  originalRow?: number;
  sourceComments?: Record<string, string>;
}

export interface Catalog {
  categories: string[];
  products: Product[];
}
