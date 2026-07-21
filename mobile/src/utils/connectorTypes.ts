// utils/connectorTypes.ts
import { ConnectorTypeCode } from '../types/ocpp';

export interface ConnectorDisplayInfo {
  code: ConnectorTypeCode;
  label: string;
  icon: string;
  color: string; // Tailwind bg-* class
  isDC: boolean;
}

const DC_TYPES: ConnectorTypeCode[] = ['CCS1', 'CCS2', 'CHAdeMO', 'GBTDC'];

export const CONNECTOR_INFO: Record<ConnectorTypeCode, ConnectorDisplayInfo> = {
  CCS2: {
    code: 'CCS2',
    label: 'CCS Combo 2',
    icon: '⚡',
    color: 'bg-emerald-100 text-emerald-800',
    isDC: true,
  },
  CCS1: {
    code: 'CCS1',
    label: 'CCS Combo 1',
    icon: '⚡',
    color: 'bg-emerald-100 text-emerald-800',
    isDC: true,
  },
  CHAdeMO: {
    code: 'CHAdeMO',
    label: 'CHAdeMO',
    icon: '🔶',
    color: 'bg-orange-100 text-orange-800',
    isDC: true,
  },
  GBTDC: {
    code: 'GBTDC',
    label: 'GB/T DC',
    icon: '🔷',
    color: 'bg-orange-100 text-orange-800',
    isDC: true,
  },
  Type2: {
    code: 'Type2',
    label: 'Type 2 (Mennekes)',
    icon: '🔹',
    color: 'bg-blue-100 text-blue-800',
    isDC: false,
  },
  Type1: {
    code: 'Type1',
    label: 'Type 1 (J1772)',
    icon: '🔸',
    color: 'bg-indigo-100 text-indigo-800',
    isDC: false,
  },
  Schuko: {
    code: 'Schuko',
    label: 'Household (Schuko)',
    icon: '🏠',
    color: 'bg-gray-100 text-gray-700',
    isDC: false,
  },
  Unknown: {
    code: 'Unknown',
    label: 'Connector',
    icon: '🔌',
    color: 'bg-gray-200 text-gray-600',
    isDC: false,
  },
};

export const getConnectorDisplay = (type?: string | null): ConnectorDisplayInfo => {
  if (!type) return CONNECTOR_INFO['Unknown'];
  const normalized = Object.keys(CONNECTOR_INFO).find(
    key => key.toLowerCase() === type.toLowerCase()
  ) as ConnectorTypeCode | undefined;
  return normalized ? CONNECTOR_INFO[normalized] : CONNECTOR_INFO['Unknown'];
};

export const isDcConnector = (type?: string | null): boolean => {
  const info = getConnectorDisplay(type);
  return info.isDC;
};