import { default as React } from 'react';
import { ChartType, DefaultDataPoint, ChartData, ChartOptions, Plugin, UpdateMode } from 'chart.js';
import { FloatElementItem } from '../../containers/store';
export interface ChartProps<TType extends ChartType = ChartType, TData = DefaultDataPoint<TType>, TLabel = unknown> {
    width: number;
    height: number;
    uuid?: string;
    type: ChartType;
    data: ChartData<TType, TData, TLabel>;
    options?: ChartOptions<TType>;
    plugins?: Plugin<TType>[];
    redraw?: boolean;
    updateMode?: UpdateMode;
}
declare const Chart: React.FunctionComponent<FloatElementItem>;
export default Chart;
