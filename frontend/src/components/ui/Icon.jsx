// Material Symbols icon component
// Usage: <Icon name="inventory_2" size="md" weight="regular" />
export default function Icon({
  name,
  size = 'md',
  weight = 'regular',
  fill = 'none',
  color,
  className = '',
  ...props
}) {
  const sizeMap = { xs: 16, sm: 20, md: 24, lg: 28, xl: 32, '2xl': 40 };
  const weightMap = { light: 300, regular: 400, medium: 500, semibold: 600, bold: 700 };
  const fillMap = { none: 0, subtle: 0.25, half: 0.5, full: 1 };
  const gradeMap = { normal: 0, dark: 100, light: -25 };

  const px = sizeMap[size] ?? 24;

  return (
    <span
      className={`msi ${className}`}
      style={{
        fontSize: px,
        color: color,
        fontVariationSettings: `'wght' ${weightMap[weight] ?? 400}, 'FILL' ${fillMap[fill] ?? 0}, 'GRAD' ${gradeMap.normal}, 'opsz' ${px}`,
      }}
      aria-hidden="true"
      {...props}
    >
      {name}
    </span>
  );
}
