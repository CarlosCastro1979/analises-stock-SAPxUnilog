import { ActivityIndicator, StyleSheet, View } from 'react-native';
import { useHousehold } from '../context/HouseholdContext';
import { HouseholdCodeScreen } from '../screens/HouseholdCodeScreen';
import { HomeScreen } from '../screens/HomeScreen';
import { colors } from '../theme';

export default function Index() {
  const { ready, householdCode } = useHousehold();

  if (!ready) {
    return (
      <View style={styles.boot}>
        <ActivityIndicator color="#fff" size="large" />
      </View>
    );
  }

  if (!householdCode) {
    return <HouseholdCodeScreen />;
  }

  return <HomeScreen />;
}

const styles = StyleSheet.create({
  boot: {
    flex: 1,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
  },
});
