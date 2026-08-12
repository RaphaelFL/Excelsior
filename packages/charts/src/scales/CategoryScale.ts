export class CategoryScale {
  constructor(
    private readonly categories: readonly string[],
    private readonly range: [number, number]
  ) {}

  map(index: number): number {
    const [rangeMin, rangeMax] = this.range;
    const count = this.categories.length;
    if (count <= 1) {
      return (rangeMin + rangeMax) / 2;
    }
    const clampedIndex = Math.max(0, Math.min(count - 1, Math.round(index)));
    const step = (rangeMax - rangeMin) / count;
    return rangeMin + step * (clampedIndex + 0.5);
  }

  label(index: number): string {
    return this.categories[index] ?? String(index);
  }
}
