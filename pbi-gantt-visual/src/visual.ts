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

const SVG_NS: string = "http://www.w3.org/2000/svg";

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly root: HTMLDivElement;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.root = document.createElement("div");
        this.root.className = "ganttRoot";
        options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        while (this.root.firstChild) {
            this.root.removeChild(this.root.firstChild);
        }

        const dataView: DataView | undefined = options.dataViews?.[0];
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

        const viewportWidth: number = Math.max(options.viewport.width, 420);
        const rowHeight: number = 30;
        const timelinePadding: number = 16;
        const chartTop: number = 48;
        const leftWidth: number = this.calculateLeftWidth(tasks);
        const timelineWidth: number = Math.max(280, viewportWidth - leftWidth - timelinePadding * 2);
        const svgWidth: number = leftWidth + timelineWidth + timelinePadding * 2;
        const svgHeight: number = chartTop + tasks.length * rowHeight + 20;

        const svg: SVGSVGElement = this.svgEl("svg") as SVGSVGElement;
        svg.setAttribute("class", "ganttSvg");
        svg.setAttribute("width", String(svgWidth));
        svg.setAttribute("height", String(svgHeight));

        this.drawHeader(svg, leftWidth, timelinePadding, timelineWidth, minDate, maxDate);

        tasks.forEach((task: TaskRow, index: number) => {
            this.drawTaskRow(
                svg,
                task,
                index,
                chartTop,
                rowHeight,
                leftWidth,
                timelinePadding,
                timelineWidth,
                minDate,
                maxDate
            );
        });

        this.root.appendChild(svg);
    }

    private renderEmpty(message: string): void {
        const empty: HTMLDivElement = document.createElement("div");
        empty.className = "ganttEmpty";
        empty.textContent = message;
        this.root.appendChild(empty);
    }

    private drawHeader(
        svg: SVGSVGElement,
        leftWidth: number,
        timelinePadding: number,
        timelineWidth: number,
        minDate: Date,
        maxDate: Date
    ): void {
        const taskHeader: SVGTextElement = this.svgEl("text") as SVGTextElement;
        taskHeader.setAttribute("x", String(12));
        taskHeader.setAttribute("y", String(25));
        taskHeader.setAttribute("class", "ganttHeaderLabel");
        taskHeader.textContent = "Task / Groups";
        svg.appendChild(taskHeader);

        const axisStartX: number = leftWidth + timelinePadding;
        const axisY: number = 30;
        const tickCount: number = 6;
        const totalMs: number = Math.max(1, maxDate.getTime() - minDate.getTime());

        for (let i: number = 0; i <= tickCount; i += 1) {
            const ratio: number = i / tickCount;
            const x: number = axisStartX + ratio * timelineWidth;
            const tickDate: Date = new Date(minDate.getTime() + totalMs * ratio);

            const gridLine: SVGLineElement = this.svgEl("line") as SVGLineElement;
            gridLine.setAttribute("x1", String(x));
            gridLine.setAttribute("x2", String(x));
            gridLine.setAttribute("y1", String(axisY + 8));
            gridLine.setAttribute("y2", String(10000));
            gridLine.setAttribute("class", "ganttGridLine");
            svg.appendChild(gridLine);

            const label: SVGTextElement = this.svgEl("text") as SVGTextElement;
            label.setAttribute("x", String(x));
            label.setAttribute("y", String(axisY));
            label.setAttribute("text-anchor", "middle");
            label.setAttribute("class", "ganttAxisLabel");
            label.textContent = this.formatDate(tickDate);
            svg.appendChild(label);
        }
    }

    private drawTaskRow(
        svg: SVGSVGElement,
        task: TaskRow,
        rowIndex: number,
        chartTop: number,
        rowHeight: number,
        leftWidth: number,
        timelinePadding: number,
        timelineWidth: number,
        minDate: Date,
        maxDate: Date
    ): void {
        const y: number = chartTop + rowIndex * rowHeight;
        const centerY: number = y + rowHeight / 2;

        const taskLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        taskLabel.setAttribute("x", String(12));
        taskLabel.setAttribute("y", String(centerY));
        taskLabel.setAttribute("class", "ganttTaskLabel");
        taskLabel.textContent = this.composeTaskLabel(task);
        if (task.url) {
            taskLabel.style.cursor = "pointer";
            taskLabel.addEventListener("click", () => this.host.launchUrl(task.url));
        }
        svg.appendChild(taskLabel);

        const statusLabel: SVGTextElement = this.svgEl("text") as SVGTextElement;
        statusLabel.setAttribute("x", String(leftWidth - 8));
        statusLabel.setAttribute("y", String(centerY));
        statusLabel.setAttribute("text-anchor", "end");
        statusLabel.setAttribute("class", "ganttStatusText");
        statusLabel.textContent = task.status;
        svg.appendChild(statusLabel);

        const timelineStartX: number = leftWidth + timelinePadding;
        const barY: number = y + 8;
        const barHeight: number = 14;

        this.drawBars(
            svg,
            task,
            timelineStartX,
            barY,
            barHeight,
            timelineWidth,
            minDate,
            maxDate
        );
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
        const xForDate = (dateValue: Date): number => {
            const total: number = Math.max(1, maxDate.getTime() - minDate.getTime());
            const pct: number = (dateValue.getTime() - minDate.getTime()) / total;
            return timelineStartX + Math.max(0, Math.min(1, pct)) * timelineWidth;
        };

        if (task.plannedStartDate && task.plannedFinishDate) {
            const plannedX: number = xForDate(task.plannedStartDate);
            const plannedW: number = Math.max(2, xForDate(task.plannedFinishDate) - plannedX);

            const plannedBar: SVGRectElement = this.svgEl("rect") as SVGRectElement;
            plannedBar.setAttribute("x", String(plannedX));
            plannedBar.setAttribute("y", String(barY));
            plannedBar.setAttribute("width", String(plannedW));
            plannedBar.setAttribute("height", String(barHeight));
            plannedBar.setAttribute("rx", "3");
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
                plannedProgress.setAttribute("rx", "3");
                plannedProgress.setAttribute("class", "ganttPlannedProgress");
                svg.appendChild(plannedProgress);
            }
        }

        if (task.startDate && task.endDate) {
            const actualX: number = xForDate(task.startDate);
            const actualW: number = Math.max(2, xForDate(task.endDate) - actualX);

            const actualBar: SVGRectElement = this.svgEl("rect") as SVGRectElement;
            actualBar.setAttribute("x", String(actualX));
            actualBar.setAttribute("y", String(barY + 2));
            actualBar.setAttribute("width", String(actualW));
            actualBar.setAttribute("height", String(barHeight - 4));
            actualBar.setAttribute("rx", "3");
            actualBar.setAttribute("class", "ganttActualBar");
            actualBar.appendChild(this.makeTitle(`Actual: ${this.formatDate(task.startDate)} - ${this.formatDate(task.endDate)}`));
            svg.appendChild(actualBar);

            const actualProgressW: number = actualW * task.progressPct;
            if (actualProgressW > 0) {
                const actualProgress: SVGRectElement = this.svgEl("rect") as SVGRectElement;
                actualProgress.setAttribute("x", String(actualX));
                actualProgress.setAttribute("y", String(barY + 2));
                actualProgress.setAttribute("width", String(actualProgressW));
                actualProgress.setAttribute("height", String(barHeight - 4));
                actualProgress.setAttribute("rx", "3");
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

        tasks.forEach((task) => {
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

        return values.reduce((min: Date, current: Date) => {
            return current.getTime() < min.getTime() ? current : min;
        });
    }

    private maxDate(tasks: TaskRow[]): Date | null {
        const values: Date[] = [];

        tasks.forEach((task) => {
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

        return values.reduce((max: Date, current: Date) => {
            return current.getTime() > max.getTime() ? current : max;
        });
    }

    private composeTaskLabel(task: TaskRow): string {
        const groupParts: string[] = [task.group1, task.group2, task.group3].filter((value) => value.length > 0);
        if (groupParts.length === 0) {
            return `${task.sequence}. ${task.taskName}`;
        }

        return `${task.sequence}. ${task.taskName} [${groupParts.join(" / ")}]`;
    }

    private formatDate(dateValue: Date): string {
        return dateValue.toLocaleDateString();
    }

    private calculateLeftWidth(tasks: TaskRow[]): number {
        const maxTextLength: number = tasks
            .map((task) => this.composeTaskLabel(task).length)
            .reduce((max: number, current: number) => Math.max(max, current), 24);

        return Math.min(560, Math.max(260, maxTextLength * 7 + 30));
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
