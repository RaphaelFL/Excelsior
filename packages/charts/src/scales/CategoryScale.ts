export class CategoryScale {
  constructor(
    private readonly categories: readonly string[],
    private readonly range: [number, number]
  ) {}

  map(index: number): number {
    const [rangeMin, rangeMax] = this.range;
    if (this.categories.length <= 1) {
      return (rangeMin + rangeMax) / 2;
    }
    const step = (rangeMax - rangeMin) / (this.categories.length - 1);
    return rangeMin + step * index;
  }

  label(index: number): string {
    return this.categories[index] ?? String(index);
  }
}
