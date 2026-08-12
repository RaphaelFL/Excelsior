export interface LinearScaleOptions {
  type?: "linear" | "log";
  reverse?: boolean;
}

export class LinearScale {
  private readonly type: "linear" | "log";
  private readonly reverse: boolean;

  constructor(
    private readonly domain: [number, number],
    private readonly range: [number, number],
    options?: LinearScaleOptions
  ) {
    this.type = options?.type === "log" ? "log" : "linear";
    this.reverse = options?.reverse === true;
  }

  map(value: number): number {
    const [rawDomainMin, rawDomainMax] = this.domain;
    const [domainMin, domainMax] = this.reverse ? ([rawDomainMax, rawDomainMin] as [number, number]) : this.domain;
    const [rangeMin, rangeMax] = this.range;
    if (domainMax === domainMin) {
      return rangeMin;
    }

    const mappedValue = this.type === "log" ? toLogValue(value, domainMin) : value;
    const mappedDomainMin = this.type === "log" ? toLogValue(domainMin, domainMin) : domainMin;
    const mappedDomainMax = this.type === "log" ? toLogValue(domainMax, domainMin) : domainMax;
    const ratio = (mappedValue - mappedDomainMin) / (mappedDomainMax - mappedDomainMin || Number.EPSILON);
    return rangeMin + ratio * (rangeMax - rangeMin);
  }

  invert(pixel: number): number {
    const [rawDomainMin, rawDomainMax] = this.domain;
    const [domainMin, domainMax] = this.reverse ? ([rawDomainMax, rawDomainMin] as [number, number]) : this.domain;
    const [rangeMin, rangeMax] = this.range;
    if (rangeMax === rangeMin) {
      return domainMin;
    }

    const ratio = (pixel - rangeMin) / (rangeMax - rangeMin);
    if (this.type === "log") {
      const minLog = toLogValue(domainMin, domainMin);
      const maxLog = toLogValue(domainMax, domainMin);
      return 10 ** (minLog + ratio * (maxLog - minLog));
    }
    return domainMin + ratio * (domainMax - domainMin);
  }

  ticks(count: number): number[] {
    const [rawDomainMin, rawDomainMax] = this.domain;
    const [domainMin, domainMax] = this.reverse ? ([rawDomainMax, rawDomainMin] as [number, number]) : this.domain;
    const safeCount = Math.max(2, Math.floor(count));
    if (this.type === "log") {
      const minLog = toLogValue(domainMin, domainMin);
      const maxLog = toLogValue(domainMax, domainMin);
      const step = (maxLog - minLog) / (safeCount - 1);
      return Array.from({ length: safeCount }, (_, index) => 10 ** (minLog + step * index));
    }
    const step = (domainMax - domainMin) / (safeCount - 1);
    return Array.from({ length: safeCount }, (_, index) => domainMin + step * index);
  }
}

const toLogValue = (value: number, fallbackMin: number): number => {
  if (Number.isFinite(value) && value > 0) {
    return Math.log10(value);
  }
  const safeFallback = Number.isFinite(fallbackMin) && fallbackMin > 0 ? fallbackMin : 1;
  return Math.log10(safeFallback);
};
