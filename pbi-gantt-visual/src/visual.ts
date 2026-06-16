import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import PrimitiveValue = powerbi.PrimitiveValue;

interface TaskRow {
    sequence: number;
    taskName: string;
    startDate: Date | null;
    endDate: Date | null;
    plannedStartDate: Date | null;
    plannedFinishDate: Date | null;
    group1: string;
    group2: string;
    group3: string;
    progressPct: number;
    plannedProgressPct: number;
    status: string;
    url: string;
}

interface GroupNode {
    name: string;
    level: number;
    children: Array<GroupNode | TaskRow>;
    isGroup: boolean;
    groupPath: string[];
}

interface LayoutMetrics {
    seqColX: number;
    taskColX: number;
    progressColX: number;
    statusColX: number;
    leftWidth: number;
    timelineStartX: number;
    timelineWidth: number;
}

interface DisplayRow {
    kind: "group" | "task";
    depth: number;
    group?: GroupNode;
    task?: TaskRow;
    pathKey: string;
}

enum TimelineScale {
    Daily = "daily",
    Weekly = "weekly",
    Monthly = "monthly",
    Quarterly = "quarterly"
}

const SVG_NS: string = "http://www.w3.org/2000/svg";

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly root: HTMLDivElement;
    private collapsedGroups: Set<string> = new Set();
    private lastUpdateOptions: VisualUpdateOptions | null = null;
    private timelineScale: TimelineScale = TimelineScale.Daily;
    private timelineScaleOverride: TimelineScale | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.root = document.createElement("div");
        this.root.className = "ganttRoot";
        options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        this.lastUpdateOptions = options;
        this.render();
    }

    private render(): void {
        if (!this.lastUpdateOptions) {
            return;
        }

        while (this.root.firstChild) {
            this.root.removeChild(this.root.firstChild);
        }

        const dataView: DataView | undefined = this.lastUpdateOptions.dataViews?.[0];
        const table: DataViewTable | undefined = dataView?.table;

        if (!table || !table.columns || table.columns.length === 0 || !table.rows || table.rows.length === 0) {
            this.renderEmpty("Add fields to build the Gantt chart.");
            return;
        }

        const tasks: TaskRow[] = this.convertTable(table);
        if (tasks.length === 0) {
            this.renderEmpty("No valid rows were found for the selected fields.");
            return;
        }

        tasks.sort((a: TaskRow, b: TaskRow) => a.sequence - b.sequence);

        const minDate: Date | null = this.minDate(tasks);
        const maxDate: Date | null = this.maxDate(tasks);
        if (!minDate || !maxDate) {
            this.renderEmpty("Start/finish dates are required to render tasks.");
            return;
        }

        const hierarchy: GroupNode[] = this.buildHierarchy(tasks);
        const displayRows: DisplayRow[] = this.buildDisplayRows(hierarchy);

        const daysDiff: number = (maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24);
        this.timelineScale = this.timelineScaleOverride ?? this.selectTimelineScale(daysDiff);

        const viewportWidth: number = Math.max(this.lastUpdateOptions.viewport.width, 420);
        const rowHeight: number = 26;
        const headerHeight: number = 62;
        const chartTop: number = headerHeight;
        const layout: LayoutMetrics = this.calculateLayout(hierarchy, viewportWidth);
        const svgWidth: number = layout.leftWidth + layout.timelineWidth;
        const svgHeight: number = chartTop + displayRows.length * rowHeight + 18;

        const svg: SVGSVGElement = this.svgEl("svg") as SVGSVGElement;
        svg.setAttribute("class", "ganttSvg");
        svg.setAttribute("width", String(svgWidth));
        svg.setAttribute("height", String(svgHeight));

        this.drawHeader(svg, layout, minDate, maxDate, chartTop);

        displayRows.forEach((row: DisplayRow, rowIndex: number) => {
            this.drawRowBackground(svg, rowIndex, chartTop, rowHeight, layout);
            if (row.kind === "group" && row.group) {
                this.drawGroupRow(svg, row.group, row.depth, rowIndex, chartTop, rowHeight, layout, row.pathKey);
            }
            if (row.kind === "task" && row.task) {
                this.drawTaskRow(svg, row.task, row.depth, rowIndex, chartTop, rowHeight, layout, minDate, maxDate);
            }
        });

        this.root.appendChild(svg);
    }

    private renderEmpty(message: string): void {
        const empty: HTMLDivElement = document.createElement("div");
        empty.className = "ganttEmpty";
        empty.textContent = message;
        this.root.appendChild(empty);
    }

    private buildDisplayRows(hierarchy: GroupNode[]): DisplayRow[] {
        const rows: DisplayRow[] = [];

        const visit = (nodes: Array<GroupNode | TaskRow>, depth: number, parentPath: string[]): void => {
            nodes.forEach((node) => {
                if (this.isGroupNode(node)) {
                    const group: GroupNode = node as GroupNode;
                    const path: string[] = [...parentPath, group.name];
                    const pathKey: string = path.join(" / ");

                    rows.push({ kind: "group", depth, group, pathKey });

                    if (!this.collapsedGroups.has(pathKey)) {
                        visit(group.children, depth + 1, path);
                    }
                } else {
                    const task: TaskRow = node as TaskRow;
                    rows.push({ kind: "task", depth, task, pathKey: parentPath.join(" / ") });
                }
            });
        };

        visit(hierarchy, 0, []);
        return rows;
    }

    private buildHierarchy(tasks: TaskRow[]): GroupNode[] {
        interface TempGroupNode {
            name: string;
            level: number;
            children: Map<string, TempGroupNode | TaskRow>;
            isGroup: boolean;
            groupPath: string[];
        }

        const root: Map<string, TempGroupNode | TaskRow> = new Map();

        tasks.forEach((task: TaskRow) => {
            const groups: string[] = [task.group1, task.group2, task.group3].filter((g) => g.length > 0);

            if (groups.length === 0) {
                root.set(`task_${task.sequence}_${task.taskName}`, task);
                return;
            }

            let current: Map<string, TempGroupNode | TaskRow> = root;
            const path: string[] = [];

            groups.forEach((group: string, level: number) => {
                path.push(group);
                const pathKey: string = path.join(" / ");
                const existing: TempGroupNode | TaskRow | undefined = current.get(pathKey);

                let groupNode: TempGroupNode;
                if (!existing || !(existing as TempGroupNode).isGroup) {
                    groupNode = {
                        name: group,
                        level,
                        children: new Map(),
                        isGroup: true,
                        groupPath: path.slice()
                    };
                    current.set(pathKey, groupNode);
                } else {
                    groupNode = existing as TempGroupNode;
                }

                current = groupNode.children;
            });

            current.set(`task_${task.sequence}_${task.taskName}`, task);
        });

        const convert = (map: Map<string, TempGroupNode | TaskRow>): GroupNode[] => {
            const nodes: GroupNode[] = [];
            const tasksAtLevel: TaskRow[] = [];

            Array.from(map.values()).forEach((item) => {
                if ((item as TempGroupNode).isGroup) {
                    const group: TempGroupNode = item as TempGroupNode;
                    nodes.push({
                        name: group.name,
                        level: group.level,
                        isGroup: true,
                        groupPath: group.groupPath,
                        children: convert(group.children)
                    });
                } else {
                    tasksAtLevel.push(item as TaskRow);
                }
            });

            tasksAtLevel.sort((a: TaskRow, b: TaskRow) => a.sequence - b.sequence);
            nodes.sort((a: GroupNode, b: GroupNode) => a.name.localeCompare(b.name));
            return [...nodes, ...(tasksAtLevel as unknown as GroupNode[])];
        };

        return convert(root);
    }

    private drawHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        this.drawTableHeader(svg, layout, chartTop);
        this.drawScalePicker(svg, layout, chartTop);
        this.drawTimelineHeader(svg, layout, minDate, maxDate, chartTop);
    }

    private drawTableHeader(svg: SVGSVGElement, layout: LayoutMetrics, chartTop: number): void {
        const headerBg: SVGRectElement = this.svgEl("rect") as SVGRectElement;
        headerBg.setAttribute("x", "0");
        headerBg.setAttribute("y", "0");
        headerBg.setAttribute("width", String(layout.leftWidth));
        headerBg.setAttribute("height", String(chartTop));
        headerBg.setAttribute("class", "ganttHeaderBg");
        svg.appendChild(headerBg);

        const columns: Array<{ x: number; name: string }> = [
            { x: layout.seqColX + 8, name: "#" },
            { x: layout.taskColX + 8, name: "Task" },
            { x: layout.progressColX + 8, name: "%" },
            { x: layout.statusColX + 8, name: "Status" }
        ];

        columns.forEach((col) => {
            const label: SVGTextElement = this.svgEl("text") as SVGTextElement;
            label.setAttribute("x", String(col.x));
            label.setAttribute("y", "38");
            label.setAttribute("class", "ganttHeaderLabel");
            label.textContent = col.name;
            svg.appendChild(label);
        });

        [layout.taskColX, layout.progressColX, layout.statusColX, layout.leftWidth].forEach((x) => {
            const line: SVGLineElement = this.svgEl("line") as SVGLineElement;
            line.setAttribute("x1", String(x));
            line.setAttribute("x2", String(x));
            line.setAttribute("y1", "0");
            line.setAttribute("y2", "10000");
            line.setAttribute("class", "ganttTableDivider");
            svg.appendChild(line);
        });

        const baseline: SVGLineElement = this.svgEl("line") as SVGLineElement;
        baseline.setAttribute("x1", "0");
        baseline.setAttribute("x2", String(layout.leftWidth));
        baseline.setAttribute("y1", String(chartTop));
        baseline.setAttribute("y2", String(chartTop));
        baseline.setAttribute("class", "ganttTableDividerStrong");
        svg.appendChild(baseline);
    }

    private drawScalePicker(svg: SVGSVGElement, layout: LayoutMetrics, chartTop: number): void {
        const scales: TimelineScale[] = [TimelineScale.Daily, TimelineScale.Weekly, TimelineScale.Monthly, TimelineScale.Quarterly];
        const labels: Record<TimelineScale, string> = {
            [TimelineScale.Daily]: "D",
            [TimelineScale.Weekly]: "W",
            [TimelineScale.Monthly]: "M",
            [TimelineScale.Quarterly]: "Q"
        };

        const startX: number = layout.timelineStartX + 8;
        scales.forEach((scale, index) => {
            const x: number = startX + index * 22;
            const active: boolean = this.timelineScale === scale;

            const pill: SVGRectElement = this.svgEl("rect") as SVGRectElement;
            pill.setAttribute("x", String(x));
            pill.setAttribute("y", "8");
            pill.setAttribute("width", "18");
            pill.setAttribute("height", "14");
            pill.setAttribute("rx", "3");
            pill.setAttribute("class", active ? "ganttScalePill ganttScalePillActive" : "ganttScalePill");
            pill.style.cursor = "pointer";
            pill.addEventListener("click", () => {
                this.timelineScaleOverride = scale;
                this.render();
            });
            svg.appendChild(pill);

            const text: SVGTextElement = this.svgEl("text") as SVGTextElement;
            text.setAttribute("x", String(x + 9));
            text.setAttribute("y", "18");
            text.setAttribute("text-anchor", "middle");
            text.setAttribute("class", active ? "ganttScaleText ganttScaleTextActive" : "ganttScaleText");
            text.textContent = labels[scale];
            text.style.cursor = "pointer";
            text.addEventListener("click", () => {
                this.timelineScaleOverride = scale;
                this.render();
            });
            svg.appendChild(text);
        });

        const resetText: SVGTextElement = this.svgEl("text") as SVGTextElement;
        resetText.setAttribute("x", String(startX + 96));
        resetText.setAttribute("y", "18");
        resetText.setAttribute("class", "ganttAutoScaleText");
        resetText.textContent = "Auto";
        resetText.style.cursor = "pointer";
        resetText.addEventListener("click", () => {
            this.timelineScaleOverride = null;
            this.render();
        });
        svg.appendChild(resetText);

        const baseline: SVGLineElement = this.svgEl("line") as SVGLineElement;
        baseline.setAttribute("x1", String(layout.timelineStartX));
        baseline.setAttribute("x2", String(layout.timelineStartX + layout.timelineWidth));
        baseline.setAttribute("y1", String(chartTop));
        baseline.setAttribute("y2", String(chartTop));
        baseline.setAttribute("class", "ganttTableDividerStrong");
        svg.appendChild(baseline);
    }

    private drawTimelineHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        if (this.timelineScale === TimelineScale.Weekly) {
            this.drawWeeklyTimelineHeader(svg, layout, minDate, maxDate, chartTop);
            return;
        }

        if (this.timelineScale === TimelineScale.Monthly) {
            this.drawMonthlyTimelineHeader(svg, layout, minDate, maxDate, chartTop);
            return;
        }

        if (this.timelineScale === TimelineScale.Quarterly) {
            this.drawQuarterlyTimelineHeader(svg, layout, minDate, maxDate, chartTop);
            return;
        }

        this.drawDailyTimelineHeader(svg, layout, minDate, maxDate, chartTop);
    }

    private drawDailyTimelineHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        const tickCount: number = 10;
        for (let i: number = 0; i <= tickCount; i += 1) {
            const ratio: number = i / tickCount;
            const x: number = layout.timelineStartX + ratio * layout.timelineWidth;
            const tickDate: Date = new Date(minDate.getTime() + (maxDate.getTime() - minDate.getTime()) * ratio);

            const label: SVGTextElement = this.svgEl("text") as SVGTextElement;
            label.setAttribute("x", String(x));
            label.setAttribute("y", "38");
            label.setAttribute("text-anchor", "middle");
            label.setAttribute("class", "ganttAxisLabel");
            label.textContent = `${tickDate.getMonth() + 1}/${tickDate.getDate()}`;
            svg.appendChild(label);

            const line: SVGLineElement = this.svgEl("line") as SVGLineElement;
            line.setAttribute("x1", String(x));
            line.setAttribute("x2", String(x));
            line.setAttribute("y1", String(chartTop));
            line.setAttribute("y2", "10000");
            line.setAttribute("class", "ganttGridLine");
            svg.appendChild(line);
        }
    }

    private drawWeeklyTimelineHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        const dayNames: string[] = ["S", "M", "T", "W", "T", "F", "S"];
        const start: Date = new Date(minDate);
        start.setHours(0, 0, 0, 0);
        const end: Date = new Date(maxDate);
        end.setHours(0, 0, 0, 0);

        const cursor: Date = new Date(start);
        while (cursor.getTime() <= end.getTime()) {
            const next: Date = new Date(cursor);
            next.setDate(next.getDate() + 1);
            const x: number = this.xForDate(cursor, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);
            const nextX: number = this.xForDate(next, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);
            const w: number = Math.max(1, nextX - x);
            const day: number = cursor.getDay();

            if (day === 0 || day === 6) {
                const weekend: SVGRectElement = this.svgEl("rect") as SVGRectElement;
                weekend.setAttribute("x", String(x));
                weekend.setAttribute("y", String(chartTop));
                weekend.setAttribute("width", String(w));
                weekend.setAttribute("height", "10000");
                weekend.setAttribute("class", "ganttWeekendShade");
                svg.appendChild(weekend);
            }

            const dayLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
            dayLabel.setAttribute("x", String(x + w / 2));
            dayLabel.setAttribute("y", "38");
            dayLabel.setAttribute("text-anchor", "middle");
            dayLabel.setAttribute("class", "ganttAxisLabel ganttWeekdayLabel");
            dayLabel.textContent = dayNames[day];
            svg.appendChild(dayLabel);

            const dayLine: SVGLineElement = this.svgEl("line") as SVGLineElement;
            dayLine.setAttribute("x1", String(x));
            dayLine.setAttribute("x2", String(x));
            dayLine.setAttribute("y1", String(chartTop));
            dayLine.setAttribute("y2", "10000");
            dayLine.setAttribute("class", day === 1 ? "ganttWeekSeparator" : "ganttGridLine ganttSkinnyGridLine");
            svg.appendChild(dayLine);

            if (day === 1) {
                const weekLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
                weekLabel.setAttribute("x", String(x + 2));
                weekLabel.setAttribute("y", "22");
                weekLabel.setAttribute("class", "ganttWeekLabel");
                weekLabel.textContent = `W${this.getWeekNumber(cursor)}`;
                svg.appendChild(weekLabel);
            }

            cursor.setDate(cursor.getDate() + 1);
        }
    }

    private drawMonthlyTimelineHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        const monthNames: string[] = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        let cursor: Date = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
        while (cursor.getTime() <= maxDate.getTime()) {
            const next: Date = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
            const x: number = this.xForDate(cursor, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);
            const nextX: number = this.xForDate(next, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);

            const label: SVGTextElement = this.svgEl("text") as SVGTextElement;
            label.setAttribute("x", String((x + nextX) / 2));
            label.setAttribute("y", "38");
            label.setAttribute("text-anchor", "middle");
            label.setAttribute("class", "ganttAxisLabel");
            label.textContent = `${monthNames[cursor.getMonth()]} ${cursor.getFullYear()}`;
            svg.appendChild(label);

            const line: SVGLineElement = this.svgEl("line") as SVGLineElement;
            line.setAttribute("x1", String(x));
            line.setAttribute("x2", String(x));
            line.setAttribute("y1", String(chartTop));
            line.setAttribute("y2", "10000");
            line.setAttribute("class", "ganttGridLine");
            svg.appendChild(line);

            cursor = next;
        }
    }

    private drawQuarterlyTimelineHeader(
        svg: SVGSVGElement,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date,
        chartTop: number
    ): void {
        let cursor: Date = new Date(minDate.getFullYear(), Math.floor(minDate.getMonth() / 3) * 3, 1);
        while (cursor.getTime() <= maxDate.getTime()) {
            const next: Date = new Date(cursor.getFullYear(), cursor.getMonth() + 3, 1);
            const x: number = this.xForDate(cursor, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);
            const nextX: number = this.xForDate(next, minDate, maxDate, layout.timelineStartX, layout.timelineWidth);
            const q: number = Math.floor(cursor.getMonth() / 3) + 1;

            const label: SVGTextElement = this.svgEl("text") as SVGTextElement;
            label.setAttribute("x", String((x + nextX) / 2));
            label.setAttribute("y", "38");
            label.setAttribute("text-anchor", "middle");
            label.setAttribute("class", "ganttAxisLabel");
            label.textContent = `Q${q} ${cursor.getFullYear()}`;
            svg.appendChild(label);

            const line: SVGLineElement = this.svgEl("line") as SVGLineElement;
            line.setAttribute("x1", String(x));
            line.setAttribute("x2", String(x));
            line.setAttribute("y1", String(chartTop));
            line.setAttribute("y2", "10000");
            line.setAttribute("class", "ganttGridLine");
            svg.appendChild(line);

            cursor = next;
        }
    }

    private drawRowBackground(
        svg: SVGSVGElement,
        rowIndex: number,
        chartTop: number,
        rowHeight: number,
        layout: LayoutMetrics
    ): void {
        const row: SVGRectElement = this.svgEl("rect") as SVGRectElement;
        row.setAttribute("x", "0");
        row.setAttribute("y", String(chartTop + rowIndex * rowHeight));
        row.setAttribute("width", String(layout.leftWidth + layout.timelineWidth));
        row.setAttribute("height", String(rowHeight));
        row.setAttribute("class", rowIndex % 2 === 0 ? "ganttRowEven" : "ganttRowOdd");
        svg.appendChild(row);

        const line: SVGLineElement = this.svgEl("line") as SVGLineElement;
        line.setAttribute("x1", "0");
        line.setAttribute("x2", String(layout.leftWidth + layout.timelineWidth));
        line.setAttribute("y1", String(chartTop + rowIndex * rowHeight));
        line.setAttribute("y2", String(chartTop + rowIndex * rowHeight));
        line.setAttribute("class", "ganttRowRule");
        svg.appendChild(line);
    }

    private drawGroupRow(
        svg: SVGSVGElement,
        group: GroupNode,
        depth: number,
        rowIndex: number,
        chartTop: number,
        rowHeight: number,
        layout: LayoutMetrics,
        pathKey: string
    ): void {
        const y: number = chartTop + rowIndex * rowHeight;
        const centerY: number = y + rowHeight / 2;
        const indent: number = depth * 14 + 8;

        const bg: SVGRectElement = this.svgEl("rect") as SVGRectElement;
        bg.setAttribute("x", String(layout.taskColX));
        bg.setAttribute("y", String(y + 1));
        bg.setAttribute("width", String(layout.progressColX - layout.taskColX));
        bg.setAttribute("height", String(rowHeight - 2));
        bg.setAttribute("class", "ganttGroupBackground");
        svg.appendChild(bg);

        const isCollapsed: boolean = this.collapsedGroups.has(pathKey);
        const buttonX: number = layout.taskColX + indent;
        const buttonY: number = centerY - 5;

        const button: SVGRectElement = this.svgEl("rect") as SVGRectElement;
        button.setAttribute("x", String(buttonX));
        button.setAttribute("y", String(buttonY));
        button.setAttribute("width", "10");
        button.setAttribute("height", "10");
        button.setAttribute("class", "ganttExpandButton");
        button.setAttribute("rx", "2");
        button.style.cursor = "pointer";
        button.addEventListener("click", () => {
            if (this.collapsedGroups.has(pathKey)) {
                this.collapsedGroups.delete(pathKey);
            } else {
                this.collapsedGroups.add(pathKey);
            }
            this.render();
        });
        svg.appendChild(button);

        const symbolText: SVGTextElement = this.svgEl("text") as SVGTextElement;
        symbolText.setAttribute("x", String(buttonX + 5));
        symbolText.setAttribute("y", String(buttonY + 8));
        symbolText.setAttribute("text-anchor", "middle");
        symbolText.setAttribute("class", "ganttExpandSymbol");
        symbolText.textContent = isCollapsed ? "+" : "-";
        symbolText.style.cursor = "pointer";
        symbolText.addEventListener("click", () => {
            if (this.collapsedGroups.has(pathKey)) {
                this.collapsedGroups.delete(pathKey);
            } else {
                this.collapsedGroups.add(pathKey);
            }
            this.render();
        });
        svg.appendChild(symbolText);

        const groupLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        groupLabel.setAttribute("x", String(buttonX + 16));
        groupLabel.setAttribute("y", String(centerY));
        groupLabel.setAttribute("class", "ganttGroupLabel");
        groupLabel.textContent = group.name;
        svg.appendChild(groupLabel);
    }

    private drawTaskRow(
        svg: SVGSVGElement,
        task: TaskRow,
        depth: number,
        rowIndex: number,
        chartTop: number,
        rowHeight: number,
        layout: LayoutMetrics,
        minDate: Date,
        maxDate: Date
    ): void {
        const y: number = chartTop + rowIndex * rowHeight;
        const centerY: number = y + rowHeight / 2;
        const indent: number = depth * 14 + 8;

        const seqLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        seqLabel.setAttribute("x", String(layout.seqColX + 8));
        seqLabel.setAttribute("y", String(centerY));
        seqLabel.setAttribute("class", "ganttTableCell ganttSeqText");
        seqLabel.textContent = String(task.sequence);
        svg.appendChild(seqLabel);

        const taskLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        taskLabel.setAttribute("x", String(layout.taskColX + indent + 14));
        taskLabel.setAttribute("y", String(centerY));
        taskLabel.setAttribute("class", "ganttTaskLabel");
        taskLabel.textContent = task.taskName;
        if (task.url) {
            taskLabel.style.cursor = "pointer";
            taskLabel.addEventListener("click", () => this.host.launchUrl(task.url));
        }
        svg.appendChild(taskLabel);

        const taskBullet: SVGCircleElement = this.svgEl("circle") as SVGCircleElement;
        taskBullet.setAttribute("cx", String(layout.taskColX + indent + 7));
        taskBullet.setAttribute("cy", String(centerY));
        taskBullet.setAttribute("r", "2.2");
        taskBullet.setAttribute("class", "ganttTaskBullet");
        svg.appendChild(taskBullet);

        const progressText: SVGTextElement = this.svgEl("text") as SVGTextElement;
        progressText.setAttribute("x", String(layout.progressColX + 8));
        progressText.setAttribute("y", String(centerY));
        progressText.setAttribute("class", "ganttTableCell");
        progressText.textContent = `${Math.round(task.progressPct * 100)}%`;
        svg.appendChild(progressText);

        const statusLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        statusLabel.setAttribute("x", String(layout.statusColX + 8));
        statusLabel.setAttribute("y", String(centerY));
        statusLabel.setAttribute("class", "ganttTableCell ganttStatusText");
        statusLabel.textContent = task.status;
        svg.appendChild(statusLabel);

        this.drawBars(svg, task, layout.timelineStartX, y + 6, 12, layout.timelineWidth, minDate, maxDate);
    }

    private drawBars(
        svg: SVGSVGElement,
        task: TaskRow,
        timelineStartX: number,
        barY: number,
        barHeight: number,
        timelineWidth: number,
        minDate: Date,
        maxDate: Date
    ): void {
        if (task.plannedStartDate && task.plannedFinishDate) {
            const plannedX: number = this.xForDate(task.plannedStartDate, minDate, maxDate, timelineStartX, timelineWidth);
            const plannedW: number = Math.max(1, this.xForDate(task.plannedFinishDate, minDate, maxDate, timelineStartX, timelineWidth) - plannedX);

            const plannedBar: SVGRectElement = this.svgEl("rect") as SVGRectElement;
            plannedBar.setAttribute("x", String(plannedX));
            plannedBar.setAttribute("y", String(barY));
            plannedBar.setAttribute("width", String(plannedW));
            plannedBar.setAttribute("height", String(barHeight));
            plannedBar.setAttribute("rx", "2");
            plannedBar.setAttribute("class", "ganttPlannedBar");
            plannedBar.appendChild(this.makeTitle(`Planned: ${this.formatDate(task.plannedStartDate)} - ${this.formatDate(task.plannedFinishDate)}`));
            svg.appendChild(plannedBar);

            const plannedProgressW: number = plannedW * task.plannedProgressPct;
            if (plannedProgressW > 0) {
                const plannedProgress: SVGRectElement = this.svgEl("rect") as SVGRectElement;
                plannedProgress.setAttribute("x", String(plannedX));
                plannedProgress.setAttribute("y", String(barY));
                plannedProgress.setAttribute("width", String(plannedProgressW));
                plannedProgress.setAttribute("height", String(barHeight));
                plannedProgress.setAttribute("rx", "2");
                plannedProgress.setAttribute("class", "ganttPlannedProgress");
                svg.appendChild(plannedProgress);
            }
        }

        if (task.startDate && task.endDate) {
            const actualX: number = this.xForDate(task.startDate, minDate, maxDate, timelineStartX, timelineWidth);
            const actualW: number = Math.max(1, this.xForDate(task.endDate, minDate, maxDate, timelineStartX, timelineWidth) - actualX);

            const actualBar: SVGRectElement = this.svgEl("rect") as SVGRectElement;
            actualBar.setAttribute("x", String(actualX));
            actualBar.setAttribute("y", String(barY + 3));
            actualBar.setAttribute("width", String(actualW));
            actualBar.setAttribute("height", String(barHeight - 5));
            actualBar.setAttribute("rx", "2");
            actualBar.setAttribute("class", "ganttActualBar");
            actualBar.appendChild(this.makeTitle(`Actual: ${this.formatDate(task.startDate)} - ${this.formatDate(task.endDate)}`));
            svg.appendChild(actualBar);

            const actualProgressW: number = actualW * task.progressPct;
            if (actualProgressW > 0) {
                const actualProgress: SVGRectElement = this.svgEl("rect") as SVGRectElement;
                actualProgress.setAttribute("x", String(actualX));
                actualProgress.setAttribute("y", String(barY + 3));
                actualProgress.setAttribute("width", String(actualProgressW));
                actualProgress.setAttribute("height", String(barHeight - 5));
                actualProgress.setAttribute("rx", "2");
                actualProgress.setAttribute("class", "ganttActualProgress");
                svg.appendChild(actualProgress);
            }
        }
    }

    private convertTable(table: DataViewTable): TaskRow[] {
        const roleIndex: Record<string, number> = {};

        table.columns.forEach((column, index) => {
            const roles: string[] = Object.keys(column.roles ?? {});
            roles.forEach((roleName) => {
                roleIndex[roleName] = index;
            });
        });

        const read = (row: PrimitiveValue[], role: string): PrimitiveValue | null => {
            const index: number | undefined = roleIndex[role];
            return index === undefined ? null : row[index];
        };

        return (table.rows || []).map((row: PrimitiveValue[]) => {
            return {
                sequence: this.readNumber(read(row, "sequence"), Number.MAX_SAFE_INTEGER),
                taskName: this.readText(read(row, "taskName"), "(Unnamed Task)"),
                startDate: this.readDate(read(row, "startDate")),
                endDate: this.readDate(read(row, "endDate")),
                plannedStartDate: this.readDate(read(row, "plannedStartDate")),
                plannedFinishDate: this.readDate(read(row, "plannedFinishDate")),
                group1: this.readText(read(row, "group1")),
                group2: this.readText(read(row, "group2")),
                group3: this.readText(read(row, "group3")),
                progressPct: this.normalizePercent(this.readNumber(read(row, "progressPct"), 0)),
                plannedProgressPct: this.normalizePercent(this.readNumber(read(row, "plannedProgressPct"), 0)),
                status: this.readText(read(row, "status")),
                url: this.readText(read(row, "url"))
            };
        });
    }

    private readText(value: PrimitiveValue | null, fallback: string = ""): string {
        if (value === null || value === undefined) {
            return fallback;
        }

        return String(value);
    }

    private readNumber(value: PrimitiveValue | null, fallback: number): number {
        if (value === null || value === undefined) {
            return fallback;
        }

        const parsed: number = Number(value);
        return Number.isFinite(parsed) ? parsed : fallback;
    }

    private readDate(value: PrimitiveValue | null): Date | null {
        if (value === null || value === undefined) {
            return null;
        }

        const dateValue: Date = value instanceof Date ? value : new Date(String(value));
        return Number.isNaN(dateValue.getTime()) ? null : dateValue;
    }

    private normalizePercent(value: number): number {
        const normalized: number = value > 1 ? value / 100 : value;
        return Math.max(0, Math.min(1, normalized));
    }

    private minDate(tasks: TaskRow[]): Date | null {
        const values: Date[] = [];
        tasks.forEach((task: TaskRow) => {
            if (task.startDate) {
                values.push(task.startDate);
            }
            if (task.plannedStartDate) {
                values.push(task.plannedStartDate);
            }
        });

        if (values.length === 0) {
            return null;
        }

        return values.reduce((min: Date, current: Date) => current.getTime() < min.getTime() ? current : min);
    }

    private maxDate(tasks: TaskRow[]): Date | null {
        const values: Date[] = [];
        tasks.forEach((task: TaskRow) => {
            if (task.endDate) {
                values.push(task.endDate);
            }
            if (task.plannedFinishDate) {
                values.push(task.plannedFinishDate);
            }
        });

        if (values.length === 0) {
            return null;
        }

        return values.reduce((max: Date, current: Date) => current.getTime() > max.getTime() ? current : max);
    }

    private formatDate(dateValue: Date): string {
        return dateValue.toLocaleDateString();
    }

    private calculateLayout(hierarchy: GroupNode[], viewportWidth: number): LayoutMetrics {
        let maxTextLength: number = 24;

        const measureNode = (node: GroupNode | TaskRow, depth: number): void => {
            if (this.isGroupNode(node)) {
                const groupNode: GroupNode = node as GroupNode;
                maxTextLength = Math.max(maxTextLength, depth * 2 + groupNode.name.length);
                groupNode.children.forEach((child: GroupNode | TaskRow) => measureNode(child, depth + 1));
                return;
            }

            const task: TaskRow = node as TaskRow;
            maxTextLength = Math.max(maxTextLength, depth * 2 + task.taskName.length);
        };

        hierarchy.forEach((node: GroupNode) => measureNode(node, 0));

        const seqColWidth: number = 34;
        const progressColWidth: number = 44;
        const statusColWidth: number = 92;
        const taskColWidth: number = Math.min(340, Math.max(180, maxTextLength * 6 + 20));
        const leftWidth: number = seqColWidth + taskColWidth + progressColWidth + statusColWidth;
        const timelineWidth: number = Math.max(240, viewportWidth - leftWidth);

        return {
            seqColX: 0,
            taskColX: seqColWidth,
            progressColX: seqColWidth + taskColWidth,
            statusColX: seqColWidth + taskColWidth + progressColWidth,
            leftWidth,
            timelineStartX: leftWidth,
            timelineWidth
        };
    }

    private isGroupNode(node: GroupNode | TaskRow): boolean {
        return (node as GroupNode).isGroup === true;
    }

    private selectTimelineScale(daysDiff: number): TimelineScale {
        if (daysDiff <= 60) {
            return TimelineScale.Daily;
        }

        if (daysDiff <= 365) {
            return TimelineScale.Weekly;
        }

        if (daysDiff <= 1460) {
            return TimelineScale.Monthly;
        }

        return TimelineScale.Quarterly;
    }

    private getWeekNumber(date: Date): number {
        const firstDayOfYear: Date = new Date(date.getFullYear(), 0, 1);
        const pastDaysOfYear: number = (date.getTime() - firstDayOfYear.getTime()) / 86400000;
        return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
    }

    private xForDate(
        dateValue: Date,
        minDate: Date,
        maxDate: Date,
        timelineStartX: number,
        timelineWidth: number
    ): number {
        const total: number = Math.max(1, maxDate.getTime() - minDate.getTime());
        const pct: number = (dateValue.getTime() - minDate.getTime()) / total;
        return timelineStartX + Math.max(0, Math.min(1, pct)) * timelineWidth;
    }

    private svgEl(name: string): SVGElement {
        return document.createElementNS(SVG_NS, name) as SVGElement;
    }

    private makeTitle(text: string): SVGTitleElement {
        const title: SVGTitleElement = this.svgEl("title") as SVGTitleElement;
        title.textContent = text;
        return title;
    }
}
