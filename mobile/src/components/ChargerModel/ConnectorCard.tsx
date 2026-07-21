// components/ChargerModal/ConnectorCard.tsx
import { View, Text } from 'react-native';
import { ConnectorCapability } from '../../types/ocpp';
import { getConnectorDisplay } from '../../utils/connectorTypes';
import { getSpeedLabel, getSpeedCategory, getSpeedColor } from '../../utils/chargingSpeed';
import { getStatusInfo } from '../../utils/ocppStatusMapper';

export const ConnectorCard = ({ connector }: { connector: ConnectorCapability }) => {
  const connectorInfo = getConnectorDisplay(connector.type);
  const speedCategory = getSpeedCategory(connector.maxPower);
  const statusInfo = getStatusInfo(connector.status);

  return (
    <View className="p-4 border border-gray-200 rounded-xl mb-3 bg-white shadow-sm">
      {/* Primary Row: Icon + Type + Power + DC/AC */}
      <View className="flex-row items-center justify-between">
        <View className="flex-row items-center gap-3">
          <Text className="text-2xl" aria-label={connectorInfo.label}>
            {connectorInfo.icon}
          </Text>
          <View>
            <Text className="font-semibold text-base">
              {connectorInfo.label}
            </Text>
            <View className="flex-row items-center gap-2 mt-0.5">
              <Text className="text-sm font-medium text-gray-900">
                {connector.maxPower} kW
              </Text>
              <Text className={`text-xs px-2 py-0.5 rounded-full ${getSpeedColor(speedCategory)}`}>
                {getSpeedLabel(connector.maxPower)}
              </Text>
              <Text className="text-xs text-gray-500">
                {connectorInfo.isDC ? 'DC' : 'AC'}
              </Text>
            </View>
          </View>
        </View>

        {/* Status Badge */}
        <View className={`px-3 py-1.5 rounded-full ${statusInfo.colorClass}`}>
          <Text className="text-xs font-medium">{statusInfo.label}</Text>
        </View>
      </View>

      {/* Secondary Row: Format + Connector ID (subtle) */}
      {(connector.format || connector.id) && (
        <View className="mt-3 pt-3 border-t border-gray-100">
          <Text className="text-xs text-gray-400">
            {[connector.format, `ID: ${connector.id}`].filter(Boolean).join(' • ')}
          </Text>
        </View>
      )}
    </View>
  );
};