// utils/chargingSpeed.ts
export type SpeedCategory = 'standard' | 'accelerated' | 'fast' | 'ultra-fast' | 'unknown';

export const getSpeedCategory = (powerKw?: number): SpeedCategory => {
  if (!powerKw || powerKw <= 0) return 'unknown';
  if (powerKw >= 150) return 'ultra-fast';
  if (powerKw >= 50) return 'fast';
  if (powerKw >= 22) return 'accelerated';
  return 'standard';
};

export const getSpeedLabel = (powerKw?: number): string => {
  const category = getSpeedCategory(powerKw);
  const labels: Record<SpeedCategory, string> = {
    'ultra-fast': 'Ultra Fast',
    'fast': 'Fast',
    'accelerated': 'Accelerated',
    'standard': 'Standard',
    'unknown': 'Speed unknown',
  };
  return labels[category];
};

export const getSpeedColor = (category: SpeedCategory): string => {
  const colors: Record<SpeedCategory, string> = {
    'ultra-fast': 'text-red-600 bg-red-50',
    'fast': 'text-orange-600 bg-orange-50',
    'accelerated': 'text-yellow-600 bg-yellow-50',
    'standard': 'text-gray-600 bg-gray-50',
    'unknown': 'text-gray-400 bg-gray-100',
  };
  return colors[category];
};