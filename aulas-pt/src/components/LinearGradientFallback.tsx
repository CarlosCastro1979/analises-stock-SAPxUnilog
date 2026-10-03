import React from 'react';
import { StyleSheet, View, type StyleProp, type ViewStyle } from 'react-native';
import { colors } from '../theme';

/** Soft layered background without an extra native gradient dependency. */
export function LinearGradientFallback({
  children,
  style,
}: {
  children: React.ReactNode;
  style?: StyleProp<ViewStyle>;
}) {
  return (
    <View style={[styles.base, style]}>
      <View style={styles.blobTop} />
      <View style={styles.blobBottom} />
      <View style={styles.content}>{children}</View>
    </View>
  );
}

const styles = StyleSheet.create({
  base: {
    flex: 1,
    backgroundColor: colors.bg,
    overflow: 'hidden',
  },
  blobTop: {
    position: 'absolute',
    top: -80,
    right: -60,
    width: 260,
    height: 260,
    borderRadius: 130,
    backgroundColor: colors.bgSoft,
    opacity: 0.9,
  },
  blobBottom: {
    position: 'absolute',
    bottom: -40,
    left: -40,
    width: 220,
    height: 220,
    borderRadius: 110,
    backgroundColor: '#0B3032',
  },
  content: {
    flex: 1,
  },
});
