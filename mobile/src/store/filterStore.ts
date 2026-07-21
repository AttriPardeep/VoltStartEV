// src/store/filterStore.ts
import { create } from 'zustand';

export interface ChargerFilters {
  availability: 'available' | 'all';
  minPower: number;
  maxDistance: number;
  maxPrice: number;
  connectorType: string[];    //  e.g. ['CCS2', 'Type2']
  vehicleType:   string;      //  '2W' | '3W' | '4W' | 'all'
}

interface FilterState {
  filters: ChargerFilters;
  showFilterSheet: boolean;
  setFilters: (f: Partial<ChargerFilters>) => void;
  resetFilters: () => void;
  toggleFilterSheet: () => void;
}

const DEFAULT: ChargerFilters = {
  availability: 'all',
  minPower: 0,
  maxDistance: 999,
  maxPrice: 999,
  connectorType: [],  // empty = all
  vehicleType:   'all',
};

export const useFilterStore = create<FilterState>((set) => ({
  filters: DEFAULT,
  showFilterSheet: false,
  setFilters: (f) => set(s => ({ filters: { ...s.filters, ...f } })),
  resetFilters: () => set({ filters: DEFAULT }),
  toggleFilterSheet: () => set(s => ({ showFilterSheet: !s.showFilterSheet })),
}));