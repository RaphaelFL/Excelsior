import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch, type PropType } from "vue";
import { createSpreadsheet, type CreateSpreadsheetOptions, type SpreadsheetInstance } from "@excelsior/vanilla";
import type { CustomCellEditor, CustomCellRenderer } from "@excelsior/renderer-dom";

const recreate = (
  host: HTMLDivElement | null,
  options: CreateSpreadsheetOptions,
  current: SpreadsheetInstance | undefined
): SpreadsheetInstance | undefined => {
  current?.destroy();
  if (!host) {
    return undefined;
  }
  return createSpreadsheet(host, options);
};

export const Spreadsheet = defineComponent({
  name: "Spreadsheet",
  props: {
    data: Array as PropType<CreateSpreadsheetOptions["data"]>,
    settings: Object as PropType<CreateSpreadsheetOptions["settings"]>,
    metadata: Object as PropType<CreateSpreadsheetOptions["metadata"]>,
    onChange: Function as PropType<CreateSpreadsheetOptions["onChange"]>,
    cellRenderers: Array as PropType<CustomCellRenderer[]>,
    cellEditors: Array as PropType<CustomCellEditor[]>,
    includeHiddenCellsInClipboard: Boolean as PropType<boolean>,
    autofill: Object as PropType<CreateSpreadsheetOptions["autofill"]>,
    localization: Object as PropType<CreateSpreadsheetOptions["localization"]>,
    renderDebounceMs: Number as PropType<CreateSpreadsheetOptions["renderDebounceMs"]>,
    chartLimits: Object as PropType<CreateSpreadsheetOptions["chartLimits"]>,
    chartPerformance: Object as PropType<CreateSpreadsheetOptions["chartPerformance"]>,
    chartInsertPreview: Boolean as PropType<CreateSpreadsheetOptions["chartInsertPreview"]>,
    rowModel: Object as PropType<CreateSpreadsheetOptions["rowModel"]>
  },
  setup(props) {
    const hostRef = ref<HTMLDivElement | null>(null);
    let instance: SpreadsheetInstance | undefined;

    const options = (): CreateSpreadsheetOptions => ({
      data: props.data,
      settings: props.settings,
      metadata: props.metadata,
      onChange: props.onChange,
      cellRenderers: props.cellRenderers,
      cellEditors: props.cellEditors,
      includeHiddenCellsInClipboard: props.includeHiddenCellsInClipboard,
      autofill: props.autofill,
      localization: props.localization,
      renderDebounceMs: props.renderDebounceMs,
      chartLimits: props.chartLimits,
      chartPerformance: props.chartPerformance,
      chartInsertPreview: props.chartInsertPreview,
      rowModel: props.rowModel
    });

    onMounted(() => {
      instance = recreate(hostRef.value, options(), instance);
    });

    watch(
      () => [
        props.data,
        props.settings,
        props.metadata,
        props.onChange,
        props.cellRenderers,
        props.cellEditors,
        props.includeHiddenCellsInClipboard,
        props.autofill,
        props.localization,
        props.renderDebounceMs,
        props.chartLimits,
        props.chartPerformance,
        props.chartInsertPreview,
        props.rowModel
      ],
      () => {
        instance = recreate(hostRef.value, options(), instance);
      }
    );

    onBeforeUnmount(() => {
      instance?.destroy();
    });

    return () => h("div", { ref: hostRef });
  }
});