// src/utils/ocppStatus.ts
export const STATUS_LABELS: Record<string, string> = {
  Available: 'Available',
  Preparing: 'Preparing to charge',
  Charging: 'Charging',
  SuspendedEV: 'Paused — vehicle full',
  SuspendedEVSE: 'Paused — charger waiting',
  Finishing: 'Finishing session',
  Reserved: 'Reserved',
  Unavailable: 'Out of service',
  Faulted: 'Fault detected',
};

export const STATUS_COLORS: Record<string, string> = {
  Available: '#22c55e',
  Preparing: '#f59e0b',
  Charging: '#3b82f6',
  SuspendedEV: '#eab308',
  SuspendedEVSE: '#f97316',
  Finishing: '#6366f1',
  Reserved: '#8b5cf6',
  Unavailable: '#6b7280',
  Faulted: '#ef4444',
};

export function getReadableStatus(status?: string): string {
  if (!status) return 'Unknown';
  return STATUS_LABELS[status] || status;
}

export function getStatusColor(status?: string): string {
  if (!status) return '#6b7280';
  return STATUS_COLORS[status] || '#6b7280';
}