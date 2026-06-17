/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

import powerbi from "powerbi-visuals-api";
import lodashMapValues from "lodash.mapvalues";
import lodashToArray from "lodash.toarray";

import DataView = powerbi.DataView;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewCategorical = powerbi.DataViewCategorical;
import DataViewValueColumns = powerbi.DataViewValueColumns;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import DataViewCategoricalColumn = powerbi.DataViewCategoricalColumn;
import PrimitiveValue = powerbi.PrimitiveValue;

import { valueFormatter as ValueFormatter } from "powerbi-visuals-utils-formattingutils";

import { converterHelper } from "powerbi-visuals-utils-dataviewutils";
import { GanttRole } from "./enums";


const extraInformationRole = GanttRole.ExtraInformation;

export class GanttColumns<T> {

    public static getGroupedValueColumns(dataView: DataView): GanttColumns<DataViewValueColumn>[] | undefined {
        const categorical: DataViewCategorical | undefined = dataView?.categorical;
        const values: DataViewValueColumns | undefined = categorical?.values;
        const grouped: DataViewValueColumnGroup[] | undefined = values?.length ? values.grouped() : undefined;

        if (values === undefined || values.length == 0) {
            return undefined;
        }

        return grouped && grouped.map(g => lodashMapValues(
            new this<DataViewValueColumn>(),
            (n: DataViewValueColumn | null, i: string) => g.values.filter(v => v.source.roles && v.source.roles[i])[0]));
    }

    public static getCategoricalValues(dataView: DataView): GanttColumns<any> | null {
        const categorical: DataViewCategorical | undefined = dataView && dataView.categorical;
        const categories: DataViewCategoricalColumn[] = categorical && categorical.categories || [];
        const values: DataViewValueColumns = categorical && categorical.values || [] as unknown as DataViewValueColumns;
        const series: PrimitiveValue[] | undefined = categorical && values.source && this.getSeriesValues(dataView);

        return (categorical && lodashMapValues(new this<any[]>(), (n: any, i: string) => {
            let columns: PrimitiveValue[] | { [x: string]: PrimitiveValue[]; } | undefined;
            (<DataViewValueColumn[]>lodashToArray(categories))
                .concat(lodashToArray(values))
                .filter(x => x.source.roles && x.source.roles[i])
                .forEach(x => {
                    if (i === extraInformationRole && x.source.roles && x.source.roles[extraInformationRole]) {
                        if (!columns) {
                            columns = {};
                        }

                        if (x.source.format) {
                            const formatter = ValueFormatter.create({ format: x.source.format });
                            (columns as { [x: string]: PrimitiveValue[] })[x.source.displayName] = x.values.map(v => formatter.format(v));
                        } else {
                            (columns as { [x: string]: PrimitiveValue[] })[x.source.displayName] = x.values;
                        }
                    } else {
                        columns = x.values;
                    }
                });

            return columns || values.source && values.source.roles && values.source.roles[i] && series;
        })) || null;
    }

    public static getSeriesValues(dataView: DataView): PrimitiveValue[] | undefined {
        return dataView && dataView.categorical && dataView.categorical.values
            && dataView.categorical.values.map(x => converterHelper.getSeriesName(x.source));
    }

    // Data Roles
    public Legend: T | null = null;
    public Task: T | null = null;
    public TaskId: T | null = null;
    public Parent: T | null = null;
    public ParentId: T | null = null;
    public StartDate: T | null = null;
    public EndDate: T | null = null;
    public Duration: T | null = null;
    public Completion: T | null = null;
    public Resource: T | null = null;
    public ExtraInformation: T | null = null;
    public Milestones: T | null = null;
}
