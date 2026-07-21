// types/ocpp.ts
export type OcppStatus =
  | 'Available'
  | 'Preparing'
  | 'Charging'
  | 'SuspendedEV'
  | 'SuspendedEVSE'
  | 'Finishing'
  | 'Reserved'
  | 'Unavailable'
  | 'Faulted'
  | 'Offline'; // Added offline state

export type SeverityLevel = 'info' | 'warning' | 'error';

export interface StatusInfo {
  label: string;
  colorClass: string;
  severity: SeverityLevel;
}

export type ConnectorTypeCode = 
  | 'Type2' | 'Type1' | 'CCS2' | 'CCS1' | 'CHAdeMO' | 'GBTDC' | 'Schuko' | 'Unknown';

export interface ConnectorCapability {
  id: string;
  type: ConnectorTypeCode;
  maxPower: number; // kW
  voltage?: number;
  amperage?: number;
  format?: 'Socket' | 'Cable';
  status: OcppStatus;
  lastHeartbeat?: string; // ISO timestamp for offline detection
}