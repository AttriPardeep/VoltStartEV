// utils/ocppStatusMapper.ts
// React Native compatible - uses hex colors, not Tailwind classes

import { OcppStatus, StatusInfo, SeverityLevel } from '../types/ocpp';

// : Hex colors matching your old STATUS_COLOR theme
export const STATUS_MAP: Record<OcppStatus, StatusInfo> = {
  Available: {
    label: 'Available',
    colorClass: '#22c55e',  // : Green 
    severity: 'info',
  },
  Preparing: {
    label: 'Preparing to charge',
    colorClass: '#f59e0b',  // : Amber
    severity: 'info',
  },
  Charging: {
    label: 'Charging',
    colorClass: '#3b82f6',  // : Blue  
    severity: 'info',
  },
  SuspendedEV: {
    label: 'Paused — vehicle full',
    colorClass: '#f97316',  // Orange
    severity: 'warning',
  },
  SuspendedEVSE: {
    label: 'Paused — charger waiting',
    colorClass: '#f97316',  // Orange
    severity: 'warning',
  },
  Finishing: {
    label: 'Completing charge',  // : Improved UX wording
    colorClass: '#a855f7',       // Purple
    severity: 'info',
  },
  Reserved: {
    label: 'Reserved',
    colorClass: '#7c3aed',  // : Purple - actionable, not disabled
    severity: 'info',
  },
  Unavailable: {
    label: 'Out of service',
    colorClass: '#ef4444',  // : Red 
    severity: 'error',
  },
  Faulted: {
    label: 'Fault detected',
    colorClass: '#ef4444',  // : Red 
    severity: 'error',
  },
  Offline: {
    label: 'Offline',  // : New offline state
    colorClass: '#6b7280',  // : Gray 
    severity: 'warning',
  },
};

export const getStatusInfo = (status: OcppStatus | string): StatusInfo => {
  const safeStatus = (status in STATUS_MAP) ? status as OcppStatus : 'Offline';
  return STATUS_MAP[safeStatus];
};

// Optional: Helper for fault severity escalation
export const getFaultSeverity = (errorCode?: string): SeverityLevel => {
  if (!errorCode) return 'error';
  const critical = ['GroundFailure', 'PowerSwitchFailure', 'OverCurrent'];
  const warning = ['OverTemperature', 'VoltageIrregular'];
  if (critical.some(c => errorCode.includes(c))) return 'error';
  if (warning.some(w => errorCode.includes(w))) return 'warning';
  return 'error';
};