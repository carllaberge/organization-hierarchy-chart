"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;

interface HierarchyNode {
    id: string;
    parentId: string;
    display: string;
    tooltip: string;
    rowIndex: number;
    highlighted: boolean;
    children: HierarchyNode[];
    x: number;
    y: number;
    width: number;
    height: number;
    depth: number;
}

interface RichTextStyle {
    fontWeight: string;
    fontStyle: string;
    textDecoration: string;
    color: string;
    fontSize: number;
    align: "left" | "center" | "right";
}

interface RichTextSegment {
    text: string;
    style: RichTextStyle;
}

export class Visual implements IVisual {

    private events: IVisualEventService;
    private target: HTMLElement;
    private svg: SVGSVGElement;

    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private zoomScale: number = 1;
    private offsetX: number = 0;
    private offsetY: number = 0;

    private isDragging: boolean = false;
    private dragStartX: number = 0;
    private dragStartY: number = 0;

    private currentNodes: HierarchyNode[] = [];
    private collapsedNodeIds: { [key: string]: boolean } = {};
    private selectedNodeIds: { [key: string]: boolean } = {};
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private selectionIdBuilder: powerbi.visuals.ISelectionIdBuilder;
    private idColumn: powerbi.DataViewCategoryColumn = null;
    private lastRefresh: Date | null = null;
    private chartTitle: string = "Organization Chart";
    private showLastUpdated: boolean = true;
    private lastUpdatedLabel: string = "Last updated:";

    /*
    * Current visual root.
    *
    * null = display the complete hierarchy.
    * When an ID is specified, this node becomes
    * the visual root and all its descendants remain visible.
    */
    private rootNodeId: string | null = null;

    /*
    * Custom context menu used for hierarchy commands.
    */
    private contextMenu: HTMLDivElement = null;

    private readonly NODE_WIDTH = 220;
    private readonly NODE_MIN_HEIGHT = 60;
    private readonly NODE_HORIZONTAL_SPACING = 40;
    private readonly NODE_VERTICAL_SPACING = 60;

    constructor(options: VisualConstructorOptions) {

        this.events = options.host.eventService;
        this.target = options.element;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = options.host.createSelectionManager();
        this.selectionIdBuilder = options.host.createSelectionIdBuilder();
        
        while (this.target.firstChild) {
            this.target.removeChild(this.target.firstChild);
        }

        this.svg = document.createElementNS(
            "http://www.w3.org/2000/svg",
            "svg"
        );

        this.svg.setAttribute("width", "100%");
        this.svg.setAttribute("height", "100%");
        this.svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");

        this.svg.style.display = "block";
        this.svg.style.width = "100%";
        this.svg.style.height = "100%";
        this.svg.style.cursor = "grab";

        this.target.appendChild(this.svg);

        this.initializeMouseControls();

        this.svg.addEventListener(
            "click",
            async (event: MouseEvent) => {

                if (event.target === this.svg) {

                    this.selectedNodeIds = {};

                    await this.selectionManager.clear();

                    this.redrawHierarchy();
                }
            }
        );

        this.svg.addEventListener(
            "contextmenu",
            (event: MouseEvent) => {

                if (event.target === this.svg) {

                    event.preventDefault();

                    this.hideHierarchyContextMenu();

                    /*
                    * Right-clicking the empty area gives
                    * only the Reset root command.
                    */
                    if (this.rootNodeId) {

                        const fakeNode =
                            this.findNodeById(
                                this.rootNodeId
                            );

                        if (fakeNode) {

                            this.showHierarchyContextMenu(
                                fakeNode,
                                event
                            );
                        }
                    }
                }
            }
        );
    }

    public update(options: VisualUpdateOptions): void {

        this.events.renderingStarted(options);

        try {

            this.clearSvg();

            if (!options.dataViews ||
                options.dataViews.length === 0 ||
                !options.dataViews[0].categorical) {

                this.showMessage("No data available");
                this.events.renderingFinished(options);
                return;
            }

            const categorical = options.dataViews[0].categorical;

            this.formattingSettings =
                this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews[0]
                );

            this.chartTitle =
                this.formattingSettings.headerCard.chartTitle.value;

            this.showLastUpdated =
                this.formattingSettings.headerCard.showLastUpdated.value;

            this.lastUpdatedLabel =
                this.formattingSettings.headerCard.lastUpdatedLabel.value;

            if (!categorical.categories ||
                categorical.categories.length === 0) {

                this.showMessage(
                    "Add ID, Parent and Display fields to the visual."
                );

                this.events.renderingFinished(options);
                return;
            }

            const categories = categorical.categories;

            let idColumn: powerbi.DataViewCategoryColumn = null;
            let parentColumn: powerbi.DataViewCategoryColumn = null;
            let displayColumn: powerbi.DataViewCategoryColumn = null;
            let tooltipColumn: powerbi.DataViewCategoryColumn = null;

            for (const column of categories) {

                const role = column.source.roles;

                if (role && role["id"]) {
                    idColumn = column;
                }

                if (role && role["parent"]) {
                    parentColumn = column;
                }

                if (role && role["display"]) {
                    displayColumn = column;
                }

                if (role && role["tooltip"]) {
                    tooltipColumn = column;
                }
            }

            this.lastRefresh = null;

            if (
                categorical.values &&
                categorical.values.length > 0
            ) {

                for (
                    const valueColumn of categorical.values
                ) {

                    const role =
                        valueColumn.source.roles;

                    if (
                        role &&
                        role["refreshDate"] &&
                        valueColumn.values &&
                        valueColumn.values.length > 0
                    ) {

                        const refreshValue =
                            valueColumn.values[0];

                        if (
                            refreshValue !== null &&
                            refreshValue !== undefined
                        ) {

                            this.lastRefresh =
                                new Date(
                                    String(refreshValue)
                                );
                        }

                        break;
                    }
                }
            }


            if (!idColumn || !parentColumn || !displayColumn) {

                this.showMessage(
                    "Please add ID, Parent and Display fields."
                );

                this.events.renderingFinished(options);
                return;
            }

            this.idColumn = idColumn;

            const nodes = this.createNodes(
                idColumn,
                parentColumn,
                displayColumn,
                tooltipColumn
            );

            this.currentNodes = nodes;

            this.buildHierarchy(nodes);

            /*
            * If the current root no longer exists in the
            * incoming Power BI data, reset to the complete hierarchy.
            */
            if (
                this.rootNodeId &&
                !this.findNodeById(this.rootNodeId)
            ) {
                this.rootNodeId = null;
            }

            this.calculateLayout(nodes);

            this.fitToView();

            this.renderHierarchy(nodes);

            this.drawLastRefresh();

            this.events.renderingFinished(options);

        } catch (error) {

            console.error("Hierarchy chart error:", error);

            this.showMessage(
                "Error rendering hierarchy: " + String(error)
            );

            this.events.renderingFailed(
                options,
                String(error)
            );
        }
    }

    private createNodes(
        idColumn: powerbi.DataViewCategoryColumn,
        parentColumn: powerbi.DataViewCategoryColumn,
        displayColumn: powerbi.DataViewCategoryColumn,
        tooltipColumn: powerbi.DataViewCategoryColumn
    ): HierarchyNode[] {

        const nodes: HierarchyNode[] = [];

        const rowCount = idColumn.values.length;

        for (let i = 0; i < rowCount; i++) {

            const id = this.toText(idColumn.values[i]);

            const parentId = this.toText(
                parentColumn.values[i]
            );

            const display = this.normalizeLineBreaks(
                this.toText(displayColumn.values[i])
            );

            const tooltip = tooltipColumn
                ? this.toText(tooltipColumn.values[i])
                : "";

            if (!id) {
                continue;
            }
                    
            nodes.push({
                id: id,
                parentId: parentId,
                display: display,
                tooltip: tooltip,
                rowIndex: i,
                highlighted: false,
                children: [],
                x: 0,
                y: 0,
                width: 180,
                height: 60,
                depth: 0
            });
        }

        return nodes;
    }

    private buildHierarchy(nodes: HierarchyNode[]): void {

        const nodeMap: { [key: string]: HierarchyNode } = {};

        for (const node of nodes) {
            node.children = [];
            nodeMap[node.id] = node;
        }

        for (const node of nodes) {

            if (!node.parentId ||
                node.parentId === node.id) {
                continue;
            }

            const parent = nodeMap[node.parentId];

            if (parent) {
                parent.children.push(node);
            }
        }
    }

    private calculateLayout(nodes: HierarchyNode[]): void {

        let roots: HierarchyNode[];

        if (this.rootNodeId) {

            const selectedRoot =
                nodes.find(
                    node => node.id === this.rootNodeId
                );

            if (selectedRoot) {

                roots = [selectedRoot];

            } else {

                roots = nodes.filter(
                    node =>
                        !node.parentId ||
                        !nodes.some(
                            parent =>
                                parent.id === node.parentId
                        )
                );
            }

        } else {

            roots = nodes.filter(
                node =>
                    !node.parentId ||
                    !nodes.some(
                        parent =>
                            parent.id === node.parentId
                    )
            );
        }

        const horizontalSpacing =
            Math.max(
                0,
                this.formattingSettings
                    .nodeCard
                    .horizontalSpacing
                    .value
            );

        const verticalSpacing =
            Math.max(
                20,
                this.formattingSettings
                    .nodeCard
                    .verticalSpacing
                    .value
            );

        /*
        * Calculate width and height of every node.
        */
        for (const node of nodes) {

            const preliminaryLines =
                this.getLines(node.display);

            node.width =
                this.calculateNodeWidth(
                    preliminaryLines
                );

            const lines =
                this.getWrappedFormattedLines(
                    node.display,
                    node.width - 30
                );

            let totalTextHeight = 0;

            for (const line of lines) {
                totalTextHeight +=
                    this.getFormattedLineHeight(line);
            }

            const verticalPadding = 30;

            node.height =
                Math.max(
                    this.NODE_MIN_HEIGHT,
                    verticalPadding + totalTextHeight
                );
        }

        /*
        * Calculate depth and maximum height
        * for each hierarchy level.
        */
        const levelHeights: {
            [depth: number]: number;
        } = {};

        const calculateDepth =
            (
                node: HierarchyNode,
                depth: number
            ): void => {

                node.depth = depth;

                if (
                    !levelHeights[depth] ||
                    node.height > levelHeights[depth]
                ) {
                    levelHeights[depth] =
                        node.height;
                }

                for (const child of node.children) {

                    calculateDepth(
                        child,
                        depth + 1
                    );
                }
            };

        for (const root of roots) {

            calculateDepth(
                root,
                0
            );
        }

        /*
        * Calculate Y position of each level.
        *
        * verticalSpacing is applied directly
        * between levels.
        */
        const levelY: {
            [depth: number]: number;
        } = {};

        let currentY = 40;

        const maxDepth =
            nodes.length > 0
                ? Math.max(
                    ...nodes.map(
                        node => node.depth
                    )
                )
                : 0;

        for (
            let depth = 0;
            depth <= maxDepth;
            depth++
        ) {

            levelY[depth] = currentY;

            currentY +=
                (levelHeights[depth] || this.NODE_MIN_HEIGHT) +
                verticalSpacing;
        }

        /*
        * Position nodes horizontally.
        *
        * Each leaf receives horizontalSpacing
        * from the previous leaf.
        */
        let nextX = 40;

        const calculateNode =
            (
                node: HierarchyNode
            ): void => {

                node.y =
                    levelY[node.depth];

                /*
                * Leaf node.
                */
                if (
                    node.children.length === 0
                ) {

                    node.x = nextX;

                    nextX +=
                        node.width +
                        horizontalSpacing;

                    return;
                }

                /*
                * Calculate children first.
                */
                for (
                    const child of node.children
                ) {

                    calculateNode(child);
                }

                /*
                * Center parent over children.
                */
                const firstChild =
                    node.children[0];

                const lastChild =
                    node.children[
                        node.children.length - 1
                    ];

                const firstCenter =
                    firstChild.x +
                    firstChild.width / 2;

                const lastCenter =
                    lastChild.x +
                    lastChild.width / 2;

                node.x =
                    (
                        (firstCenter + lastCenter) / 2
                    ) -
                    node.width / 2;
            };

        for (const root of roots) {

            calculateNode(root);
        }
    }

    private renderHierarchy(nodes: HierarchyNode[]): void {

        const visibleNodes = this.getVisibleNodes(nodes);

        const rootGroup = document.createElementNS(
            "http://www.w3.org/2000/svg",
            "g"
        );

        rootGroup.setAttribute(
            "transform",
            "translate(" +
            this.offsetX +
            "," +
            this.offsetY +
            ") scale(" +
            this.zoomScale +
            ")"
        );

        this.svg.appendChild(rootGroup);

        // Draw connectors first
        for (const node of visibleNodes) {

            for (const child of node.children) {

                if (!this.isNodeVisible(child)) {
                    continue;
                }

                this.drawConnector(
                    rootGroup,
                    node,
                    child
                );
            }
        }

        // Draw nodes on top of connectors
        for (const node of visibleNodes) {

            this.drawNode(
                rootGroup,
                node
            );
        }
    }

    private isNodeVisible(
        node: HierarchyNode
    ): boolean {

        /*
        * If a visual root is defined,
        * the selected root and only its descendants
        * are allowed to be displayed.
        */
        if (this.rootNodeId) {

            const root =
                this.findNodeById(
                    this.rootNodeId
                );

            if (!root) {
                return false;
            }

            let current: HierarchyNode = node;

            let belongsToRoot = false;

            while (current) {

                if (current.id === root.id) {

                    belongsToRoot = true;
                    break;
                }

                if (!current.parentId) {
                    break;
                }

                current =
                    this.findNodeById(
                        current.parentId
                    );
            }

            if (!belongsToRoot) {
                return false;
            }
        }

        /*
        * Check collapsed parents.
        */
        let current: HierarchyNode = node;

        while (current.parentId) {

            const parent =
                this.findNodeById(
                    current.parentId
                );

            if (!parent) {
                break;
            }

            /*
            * The selected visual root itself must remain visible.
            */
            if (
                this.rootNodeId &&
                parent.id === this.rootNodeId
            ) {
                break;
            }

            if (
                this.collapsedNodeIds[parent.id]
            ) {
                return false;
            }

            current = parent;
        }

        return true;
    }

    private getVisibleNodes(
        nodes: HierarchyNode[]
    ): HierarchyNode[] {

        return nodes.filter(
            node => this.isNodeVisible(node)
        );
    }

    private findNodeById(
        id: string
    ): HierarchyNode {

        if (!this.currentNodes) {
            return null;
        }

        return this.currentNodes.find(
            node => node.id === id
        ) || null;
    }

    private createSelectionId(
        rowIndex: number
    ): powerbi.visuals.ISelectionId {

        return this.selectionIdBuilder
            .withCategory(
                this.idColumn,
                rowIndex
            )
            .createSelectionId();
    }
    private redrawHierarchy(): void {

        if (!this.currentNodes ||
            this.currentNodes.length === 0) {
            return;
        }

        this.clearSvg();

        this.calculateLayout(
            this.currentNodes
        );

        this.renderHierarchy(
            this.currentNodes
        );

        this.drawLastRefresh();
    }

    private drawConnector(
        parentGroup: SVGGElement,
        parent: HierarchyNode,
        child: HierarchyNode
    ): void {

        const parentX =
            parent.x + parent.width / 2;

        const parentY =
            parent.y + parent.height;

        const childX =
            child.x + child.width / 2;

        const childY =
            child.y;

        const middleY =
            parentY + ((childY - parentY) / 2);

        const path =
            document.createElementNS(
                "http://www.w3.org/2000/svg",
                "path"
            );

        const d =
            "M " + parentX + " " + parentY +
            " L " + parentX + " " + middleY +
            " L " + childX + " " + middleY +
            " L " + childX + " " + childY;

        path.setAttribute(
            "d",
            d
        );

        path.setAttribute(
            "fill",
            "none"
        );

        path.setAttribute(
            "stroke",
            this.formattingSettings
                .nodeCard
                .connectionColor
                .value
                .value
        );

        path.setAttribute(
            "stroke-width",
            this.formattingSettings
                .nodeCard
                .connectionWidth
                .value
                .toString()
        );

        const connectionStyle =
            this.formattingSettings
                .nodeCard
                .connectionStyle
                .value
                .value;

        if (connectionStyle === "dashed") {

            path.setAttribute(
                "stroke-dasharray",
                "8,5"
            );

        } else if (connectionStyle === "dotted") {

            path.setAttribute(
                "stroke-dasharray",
                "2,5"
            );
        }

        parentGroup.appendChild(path);
    }

    private getNodeBackgroundColor(node: HierarchyNode): string {

        const settings = this.formattingSettings.nodeCard;

        /*
        * If level colors are disabled,
        * use the standard background color.
        */
        if (!settings.useLevelColors.value) {
            return settings.backgroundColor.value.value;
        }

        /*
        * Select the color according to the hierarchy depth.
        *
        * Level 0 = Root
        * Level 1 = Children of Root
        * Level 2 = Grandchildren
        * Level 3 = ...
        * Level 4+ = Level 4 color
        */
        switch (node.depth) {

            case 0:
                return settings.level0Color.value.value;

            case 1:
                return settings.level1Color.value.value;

            case 2:
                return settings.level2Color.value.value;

            case 3:
                return settings.level3Color.value.value;

            case 4:
            default:
                return settings.level4Color.value.value;
        }
    }

    private setAsRoot(
        node: HierarchyNode
    ): void {

        if (!node) {
            return;
        }

        /*
        * Make the selected node the new visual root.
        */
        this.rootNodeId = node.id;

        /*
        * Recalculate the hierarchy.
        */
        this.calculateLayout(
            this.currentNodes
        );

        /*
        * Fit the new visible hierarchy
        * inside the visual.
        */
        this.fitToView();

        /*
        * Redraw.
        */
        this.clearSvg();

        this.renderHierarchy(
            this.currentNodes
        );

        this.drawLastRefresh();
    }

    private resetRoot(): void {

        /*
        * Return to the original hierarchy root(s).
        */
        this.rootNodeId = null;

        /*
        * Recalculate the hierarchy.
        */
        this.calculateLayout(
            this.currentNodes
        );

        /*
        * Fit everything back into the visual.
        */
        this.fitToView();

        /*
        * Redraw.
        */
        this.clearSvg();

        this.renderHierarchy(
            this.currentNodes
        );

        this.drawLastRefresh();
    }

    private drawNode(
        parentGroup: SVGGElement,
        node: HierarchyNode
    ): void {

        const nodeBackgroundColor = this.getNodeBackgroundColor(node);

        const group = document.createElementNS(
            "http://www.w3.org/2000/svg",
            "g"
        );

        group.setAttribute(
            "transform",
            "translate(" +
            node.x +
            "," +
            node.y +
            ")"
        );

        group.style.cursor = "pointer";

        /*
        * Node selection
        */
        group.addEventListener(
            "click",
            async (event: MouseEvent) => {

                event.stopPropagation();

                const multiSelect =
                    event.ctrlKey || event.metaKey;

                const selectionId =
                    this.createSelectionId(node.rowIndex);

                await this.selectionManager.select(
                    selectionId,
                    multiSelect
                );

                if (multiSelect) {

                    if (this.selectedNodeIds[node.id]) {

                        delete this.selectedNodeIds[node.id];

                    } else {

                        this.selectedNodeIds[node.id] = true;
                    }

                } else {

                    this.selectedNodeIds = {};
                    this.selectedNodeIds[node.id] = true;
                }

                this.redrawHierarchy();
            }
        );

        /*
        * Custom hierarchy context menu
        */
        group.addEventListener(
            "contextmenu",
            (event: MouseEvent) => {

                this.showHierarchyContextMenu(
                    node,
                    event
                );
            }
        );

        /*
        * Current box shape.
        */
        const boxShape =
            this.formattingSettings.nodeCard.boxShape.value.value;

        /*
        * Shadow
        */
        if (boxShape === "ellipse") {

            const shadow =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "ellipse"
                );

            shadow.setAttribute(
                "cx",
                (node.width / 2).toString()
            );

            shadow.setAttribute(
                "cy",
                (node.height / 2 + 3).toString()
            );

            shadow.setAttribute(
                "rx",
                (node.width / 2).toString()
            );

            shadow.setAttribute(
                "ry",
                (node.height / 2).toString()
            );

            shadow.setAttribute(
                "fill",
                "#000000"
            );

            shadow.setAttribute(
                "opacity",
                "0.12"
            );

            group.appendChild(shadow);

        } else {

            const shadow =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "rect"
                );

            shadow.setAttribute(
                "width",
                node.width.toString()
            );

            shadow.setAttribute(
                "height",
                node.height.toString()
            );

            if (boxShape === "rounded") {

                shadow.setAttribute("rx", "8");
                shadow.setAttribute("ry", "8");

            } else {

                shadow.setAttribute("rx", "0");
                shadow.setAttribute("ry", "0");
            }

            shadow.setAttribute(
                "fill",
                "#000000"
            );

            shadow.setAttribute(
                "opacity",
                "0.12"
            );

            shadow.setAttribute(
                "transform",
                "translate(0, 3)"
            );

            group.appendChild(shadow);
        }

        /*
        * Main node shape.
        */
        if (boxShape === "ellipse") {

            const ellipse =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "ellipse"
                );

            ellipse.setAttribute(
                "cx",
                (node.width / 2).toString()
            );

            ellipse.setAttribute(
                "cy",
                (node.height / 2).toString()
            );

            ellipse.setAttribute(
                "rx",
                (node.width / 2).toString()
            );

            ellipse.setAttribute(
                "ry",
                (node.height / 2).toString()
            );

            ellipse.setAttribute(
                "fill",
                nodeBackgroundColor
            );

            ellipse.setAttribute(
                "stroke",
                this.formattingSettings.nodeCard.borderColor.value.value
            );

            ellipse.setAttribute(
                "stroke-width",
                this.selectedNodeIds[node.id]
                    ? "3"
                    : "1.5"
            );

            group.appendChild(ellipse);

        } else {

            const rectangle =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "rect"
                );

            rectangle.setAttribute(
                "width",
                node.width.toString()
            );

            rectangle.setAttribute(
                "height",
                node.height.toString()
            );

            if (boxShape === "rounded") {

                rectangle.setAttribute("rx", "8");
                rectangle.setAttribute("ry", "8");

            } else {

                rectangle.setAttribute("rx", "0");
                rectangle.setAttribute("ry", "0");
            }

            rectangle.setAttribute(
                "fill",
                nodeBackgroundColor
            );

            rectangle.setAttribute(
                "stroke",
                this.formattingSettings.nodeCard.borderColor.value.value
            );

            rectangle.setAttribute(
                "stroke-width",
                this.selectedNodeIds[node.id]
                    ? "3"
                    : "1.5"
            );

            group.appendChild(rectangle);
        }

        /*
        * Expand / Collapse button
        * Positioned below the node.
        */
        if (node.children.length > 0) {

            const collapsed =
                !!this.collapsedNodeIds[node.id];

            const button = document.createElementNS(
                "http://www.w3.org/2000/svg",
                "circle"
            );

            button.setAttribute(
                "cx",
                (node.width / 2).toString()
            );

            button.setAttribute(
                "cy",
                (node.height + 14).toString()
            );

            button.setAttribute(
                "r",
                "9"
            );

            button.setAttribute(
                "fill",
                nodeBackgroundColor
            );

            button.setAttribute(
                "stroke",
                this.formattingSettings.nodeCard.borderColor.value.value
            );

            button.style.cursor = "pointer";

            button.addEventListener(
                "click",
                (event: MouseEvent) => {

                    event.stopPropagation();

                    this.collapsedNodeIds[node.id] =
                        !collapsed;

                    /*
                    * Recalculate the hierarchy layout.
                    */
                    this.calculateLayout(
                        this.currentNodes
                    );

                    /*
                    * Recalculate zoom and position so the
                    * visible hierarchy fits in the visual.
                    */
                    this.fitToView();

                    /*
                    * Remove the previous drawing.
                    */
                    this.clearSvg();

                    /*
                    * Redraw using the new dimensions.
                    */
                    this.renderHierarchy(
                        this.currentNodes
                    );

                    this.drawLastRefresh();
                }
            );

            group.appendChild(button);

            const symbol =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "text"
                );

            symbol.setAttribute(
                "x",
                (node.width / 2).toString()
            );

            symbol.setAttribute(
                "y",
                (node.height + 18).toString()
            );

            symbol.setAttribute(
                "text-anchor",
                "middle"
            );

            symbol.setAttribute(
                "font-family",
                "Segoe UI, Arial, sans-serif"
            );

            symbol.setAttribute(
                "font-size",
                "12px"
            );

            symbol.setAttribute(
                "font-weight",
                "600"
            );

            symbol.setAttribute(
                "fill",
                this.formattingSettings.nodeCard.fontColor.value.value
            );

            symbol.style.pointerEvents = "none";

            symbol.textContent =
                collapsed ? "+" : "−";

            group.appendChild(symbol);
        }

        /*
        * Node text
        */
        const lines =
            this.getWrappedFormattedLines(
                node.display,
                node.width - 30
            );

        let totalTextHeight = 0;

        for (const line of lines) {
            totalTextHeight +=
                this.getFormattedLineHeight(line);
        }

        let textY =
            (node.height - totalTextHeight) / 2;

        for (let i = 0; i < lines.length; i++) {

            const line = lines[i];

            const lineHeight =
                this.getFormattedLineHeight(line);

            const text =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "text"
                );

            const alignment =
                line.length > 0
                    ? line[0].style.align
                    : "center";

            if (alignment === "left") {
                text.setAttribute("x", "15");
                text.setAttribute("text-anchor", "start");
            }
            else if (alignment === "right") {
                text.setAttribute(
                    "x",
                    (node.width - 15).toString()
                );

                text.setAttribute(
                    "text-anchor",
                    "end"
                );
            }
            else {
                text.setAttribute(
                    "x",
                    (node.width / 2).toString()
                );

                text.setAttribute(
                    "text-anchor",
                    "middle"
                );
            }

            text.setAttribute(
                "y",
                (textY + lineHeight - 4).toString()
            );

            text.setAttribute(
                "dominant-baseline",
                "alphabetic"
            );

            text.setAttribute(
                "font-family",
                "Segoe UI, Arial, sans-serif"
            );

            text.style.pointerEvents = "none";

            for (const segment of line) {

                const tspan =
                    document.createElementNS(
                        "http://www.w3.org/2000/svg",
                        "tspan"
                    );

                tspan.textContent = segment.text;

                tspan.setAttribute(
                    "font-size",
                    segment.style.fontSize + "px"
                );

                tspan.setAttribute(
                    "font-weight",
                    segment.style.fontWeight
                );

                tspan.setAttribute(
                    "font-style",
                    segment.style.fontStyle
                );

                tspan.setAttribute(
                    "text-decoration",
                    segment.style.textDecoration
                );

                tspan.setAttribute(
                    "fill",
                    segment.style.color
                );

                text.appendChild(tspan);
            }

            group.appendChild(text);

            textY += lineHeight;
        }

        /*
        * Tooltip
        */
        if (node.tooltip) {

            const title =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "title"
                );

            title.textContent = node.tooltip;

            group.appendChild(title);
        }

        parentGroup.appendChild(group);
    }

    private getWrappedLines(
        text: string,
        maxWidth: number
    ): string[] {

        const originalLines =
            this.getLines(text);

        const result: string[] = [];

        const averageCharacterWidth = 7;

        const maxCharacters =
            Math.max(
                10,
                Math.floor(
                    maxWidth /
                    averageCharacterWidth
                )
            );

        for (const originalLine of originalLines) {

            if (!originalLine) {

                result.push("");

                continue;
            }

            const words =
                originalLine.split(/\s+/);

            let currentLine = "";

            for (const word of words) {

                const testLine =
                    currentLine
                        ? currentLine + " " + word
                        : word;

                if (
                    testLine.length <=
                    maxCharacters
                ) {

                    currentLine =
                        testLine;

                } else {

                    if (currentLine) {

                        result.push(
                            currentLine
                        );
                    }

                    /*
                    * Split very long words.
                    */
                    if (
                        word.length >
                        maxCharacters
                    ) {

                        let remaining =
                            word;

                        while (
                            remaining.length >
                            maxCharacters
                        ) {

                            result.push(
                                remaining.substring(
                                    0,
                                    maxCharacters
                                )
                            );

                            remaining =
                                remaining.substring(
                                    maxCharacters
                                );
                        }

                        currentLine =
                            remaining;

                    } else {

                        currentLine =
                            word;
                    }
                }
            }

            if (currentLine) {

                result.push(
                    currentLine
                );
            }
        }

        return result;
    }

    private getDefaultTextStyle(
        lineIndex: number
    ): RichTextStyle {

        const fontSize =
            this.formattingSettings.nodeCard.fontSize.value;

        return {
            fontWeight:
                lineIndex === 0 ? "600" : "400",

            fontStyle: "normal",

            textDecoration: "none",

            color:
                this.formattingSettings
                    .nodeCard
                    .fontColor
                    .value
                    .value,

            fontSize:
                lineIndex === 0
                    ? fontSize + 2
                    : fontSize,

            align: "center"
        };
    }

    private parseFormattedLine(
        line: string,
        lineIndex: number
    ): RichTextSegment[] {

        const result: RichTextSegment[] = [];

        let currentStyle =
            this.getDefaultTextStyle(lineIndex);

        const stack: {
            tag: string;
            previousStyle: RichTextStyle;
        }[] = [];

        const tagRegex =
            /\[\/?(bold|italic|underline|color|size|align|weight)(?:=([^\]]+))?\]/gi;

        let lastIndex = 0;
        let match: RegExpExecArray;

        while ((match = tagRegex.exec(line)) !== null) {

            const textBefore =
                line.substring(
                    lastIndex,
                    match.index
                );

            if (textBefore) {
                result.push({
                    text: textBefore,
                    style: {
                        ...currentStyle
                    }
                });
            }

            const fullTag = match[0];
            const tagName =
                match[1].toLowerCase();

            const tagValue =
                match[2];

            const isClosing =
                fullTag.indexOf("[/") === 0;

            if (isClosing) {

                for (
                    let i = stack.length - 1;
                    i >= 0;
                    i--
                ) {
                    if (stack[i].tag === tagName) {

                        currentStyle =
                            stack[i].previousStyle;

                        stack.splice(i, 1);

                        break;
                    }
                }
            }
            else {

                stack.push({
                    tag: tagName,
                    previousStyle: {
                        ...currentStyle
                    }
                });

                switch (tagName) {

                    case "bold":
                        currentStyle.fontWeight = "700";
                        break;

                    case "italic":
                        currentStyle.fontStyle = "italic";
                        break;

                    case "underline":
                        currentStyle.textDecoration =
                            "underline";
                        break;

                    case "color":
                        if (
                            tagValue &&
                            /^#[0-9a-fA-F]{3,8}$/.test(
                                tagValue.trim()
                            )
                        ) {
                            currentStyle.color =
                                tagValue.trim();
                        }
                        break;

                    case "size":

                        const parsedSize =
                            parseInt(
                                tagValue,
                                10
                            );

                        if (!isNaN(parsedSize)) {

                            currentStyle.fontSize =
                                Math.max(
                                    8,
                                    Math.min(
                                        40,
                                        parsedSize
                                    )
                                );
                        }

                        break;

                    case "weight":

                        const parsedWeight =
                            parseInt(
                                tagValue,
                                10
                            );

                        if (
                            !isNaN(parsedWeight)
                        ) {

                            currentStyle.fontWeight =
                                Math.max(
                                    100,
                                    Math.min(
                                        900,
                                        parsedWeight
                                    )
                                ).toString();
                        }

                        break;

                    case "align":

                        if (
                            tagValue === "left" ||
                            tagValue === "center" ||
                            tagValue === "right"
                        ) {
                            currentStyle.align =
                                tagValue;
                        }

                        break;
                }
            }

            lastIndex =
                match.index +
                fullTag.length;
        }

        const remainingText =
            line.substring(lastIndex);

        if (remainingText) {
            result.push({
                text: remainingText,
                style: {
                    ...currentStyle
                }
            });
        }

        return result;
    }

    private getWrappedFormattedLines(
        text: string,
        maxWidth: number
    ): RichTextSegment[][] {

        const originalLines =
            this.getLines(text);

        const result: RichTextSegment[][] = [];

        for (
            let lineIndex = 0;
            lineIndex < originalLines.length;
            lineIndex++
        ) {

            const styledSegments =
                this.parseFormattedLine(
                    originalLines[lineIndex],
                    lineIndex
                );

            if (
                styledSegments.length === 0
            ) {
                result.push([]);
                continue;
            }

            let currentLine:
                RichTextSegment[] = [];

            let currentCharacters = 0;

            for (
                const segment of styledSegments
            ) {

                const words =
                    segment.text.split(/(\s+)/);

                for (
                    const word of words
                ) {

                    if (!word) {
                        continue;
                    }

                    const cleanWord =
                        word.trim();

                    if (!cleanWord) {

                        if (
                            currentLine.length > 0
                        ) {
                            currentLine.push({
                                text: word,
                                style: {
                                    ...segment.style
                                }
                            });
                        }

                        continue;
                    }

                    const averageCharacterWidth =
                        Math.max(
                            5,
                            segment.style.fontSize *
                            0.55
                        );

                    const maxCharacters =
                        Math.max(
                            8,
                            Math.floor(
                                maxWidth /
                                averageCharacterWidth
                            )
                        );

                    if (
                        currentCharacters +
                        cleanWord.length >
                        maxCharacters &&
                        currentLine.length > 0
                    ) {

                        result.push(
                            currentLine
                        );

                        currentLine = [];

                        currentCharacters = 0;
                    }

                    currentLine.push({
                        text: cleanWord,
                        style: {
                            ...segment.style
                        }
                    });

                    currentCharacters +=
                        cleanWord.length;
                }
            }

            if (
                currentLine.length > 0
            ) {
                result.push(currentLine);
            }
        }

        return result;
    }

    private getFormattedLineHeight(
        line: RichTextSegment[]
    ): number {

        if (
            !line ||
            line.length === 0
        ) {
            return 20;
        }

        let maxFontSize = 12;

        for (
            const segment of line
        ) {

            if (
                segment.style.fontSize >
                maxFontSize
            ) {
                maxFontSize =
                    segment.style.fontSize;
            }
        }

        return Math.max(
            20,
            Math.ceil(
                maxFontSize * 1.35
            )
        );
    }

    private getLines(text: string): string[] {

        if (!text) {
            return [""];
        }

        return text
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n")
            .split("\n");
    }

    private normalizeLineBreaks(text: string): string {

        if (!text) {
            return "";
        }

        return text
            .replace(/\r\n/g, "\n")
            .replace(/\r/g, "\n");
    }

    private calculateNodeWidth(
        lines: string[]
    ): number {

        const width =
            this.formattingSettings.nodeCard.nodeWidth.value;

        return Math.max(
            120,
            Math.min(
                400,
                width
            )
        );
    }

    private toText(value: any): string {

        if (value === null ||
            value === undefined) {
            return "";
        }

        return String(value);
    }

    private clearSvg(): void {

        while (this.svg.firstChild) {
            this.svg.removeChild(
                this.svg.firstChild
            );
        }
    }

    private showMessage(message: string): void {

        const text =
            document.createElementNS(
                "http://www.w3.org/2000/svg",
                "text"
            );

        text.setAttribute(
            "x",
            "20"
        );

        text.setAttribute(
            "y",
            "40"
        );

        text.setAttribute(
            "font-family",
            "Segoe UI, Arial, sans-serif"
        );

        text.setAttribute(
            "font-size",
            "14px"
        );

        text.setAttribute(
            "fill",
            "#666666"
        );

        text.textContent = message;

        this.svg.appendChild(text);
    }

    private initializeMouseControls(): void {

        this.svg.addEventListener(
            "wheel",
            (event: WheelEvent) => {

                event.preventDefault();

                const zoomFactor =
                    event.deltaY < 0
                        ? 1.1
                        : 0.9;

                this.zoomScale *= zoomFactor;

                this.zoomScale =
                    Math.max(
                        0.2,
                        Math.min(
                            3,
                            this.zoomScale
                        )
                    );

                this.redrawFromCurrentView();
            },
            { passive: false }
        );

        this.svg.addEventListener(
            "mousedown",
            (event: MouseEvent) => {

                this.isDragging = true;

                this.dragStartX =
                    event.clientX -
                    this.offsetX;

                this.dragStartY =
                    event.clientY -
                    this.offsetY;

                this.svg.style.cursor = "grabbing";
            }
        );

        window.addEventListener(
            "mousemove",
            (event: MouseEvent) => {

                if (!this.isDragging) {
                    return;
                }

                this.offsetX =
                    event.clientX -
                    this.dragStartX;

                this.offsetY =
                    event.clientY -
                    this.dragStartY;

                this.redrawFromCurrentView();
            }
        );

        window.addEventListener(
            "mouseup",
            () => {

                this.isDragging = false;

                this.svg.style.cursor = "grab";
            }
        );
    }

    private fitToView(): void {

        if (!this.currentNodes ||
            this.currentNodes.length === 0) {
            return;
        }

        const visibleNodes =
            this.getVisibleNodes(
                this.currentNodes
            );

        if (visibleNodes.length === 0) {
            return;
        }

        let minX = Infinity;
        let minY = Infinity;
        let maxX = -Infinity;
        let maxY = -Infinity;

        for (const node of visibleNodes) {

            minX =
                Math.min(
                    minX,
                    node.x
                );

            minY =
                Math.min(
                    minY,
                    node.y
                );

            maxX =
                Math.max(
                    maxX,
                    node.x + node.width
                );

            maxY =
                Math.max(
                    maxY,
                    node.y + node.height
                );
        }

        const diagramWidth =
            maxX - minX;

        const diagramHeight =
            maxY - minY;

        const availableWidth =
            this.target.clientWidth;

        const availableHeight =
            this.target.clientHeight;

        if (
            availableWidth <= 0 ||
            availableHeight <= 0
        ) {
            return;
        }

        const padding = 30;

        const scaleX =
            (availableWidth - padding * 2) /
            diagramWidth;

        const scaleY =
            (availableHeight - padding * 2) /
            diagramHeight;

        this.zoomScale =
            Math.min(
                scaleX,
                scaleY,
                1
            );

        this.zoomScale =
            Math.max(
                0.2,
                this.zoomScale
            );

        this.offsetX =
            (availableWidth -
                diagramWidth * this.zoomScale) / 2
            -
            minX * this.zoomScale;

        this.offsetY =
            (availableHeight -
                diagramHeight * this.zoomScale) / 2
            -
            minY * this.zoomScale;
    }

    private redrawFromCurrentView(): void {

        const rootGroup =
            this.svg.firstElementChild as SVGGElement;

        if (!rootGroup) {
            return;
        }

        rootGroup.setAttribute(
            "transform",
            "translate(" +
            this.offsetX +
            "," +
            this.offsetY +
            ") scale(" +
            this.zoomScale +
            ")"
        );
    }

    public destroy(): void {

        this.hideHierarchyContextMenu();

        window.removeEventListener(
            "mousemove",
            () => { }
        );

        window.removeEventListener(
            "mouseup",
            () => { }
        );
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {

        return this.formattingSettingsService.buildFormattingModel(
            this.formattingSettings
        );
    }

    private showHierarchyContextMenu(
        node: HierarchyNode,
        event: MouseEvent
    ): void {

        event.preventDefault();
        event.stopPropagation();

        /*
        * Remove an existing menu.
        */
        this.hideHierarchyContextMenu();

        const menu =
            document.createElement("div");

        this.contextMenu = menu;

        menu.style.position = "fixed";
        menu.style.left = event.clientX + "px";
        menu.style.top = event.clientY + "px";
        menu.style.background = "#ffffff";
        menu.style.border = "1px solid #cccccc";
        menu.style.borderRadius = "4px";
        menu.style.boxShadow =
            "0 2px 8px rgba(0,0,0,0.20)";
        menu.style.padding = "4px 0";
        menu.style.minWidth = "170px";
        menu.style.zIndex = "999999";
        menu.style.fontFamily =
            "Segoe UI, Arial, sans-serif";
        menu.style.fontSize = "13px";

        /*
        * Set as root.
        */
        const setRootItem =
            document.createElement("div");

        setRootItem.textContent =
            "Set as root";

        setRootItem.style.padding =
            "8px 14px";

        setRootItem.style.cursor =
            "pointer";

        setRootItem.addEventListener(
            "mouseenter",
            () => {
                setRootItem.style.background =
                    "#eeeeee";
            }
        );

        setRootItem.addEventListener(
            "mouseleave",
            () => {
                setRootItem.style.background =
                    "#ffffff";
            }
        );

        setRootItem.addEventListener(
            "click",
            () => {

                this.setAsRoot(node);

                this.hideHierarchyContextMenu();
            }
        );

        menu.appendChild(setRootItem);

        /*
        * Reset root.
        */
        const resetRootItem =
            document.createElement("div");

        resetRootItem.textContent =
            "Reset root";

        resetRootItem.style.padding =
            "8px 14px";

        resetRootItem.style.cursor =
            this.rootNodeId
                ? "pointer"
                : "default";

        resetRootItem.style.opacity =
            this.rootNodeId
                ? "1"
                : "0.45";

        resetRootItem.addEventListener(
            "mouseenter",
            () => {

                if (this.rootNodeId) {
                    resetRootItem.style.background =
                        "#eeeeee";
                }
            }
        );

        resetRootItem.addEventListener(
            "mouseleave",
            () => {
                resetRootItem.style.background =
                    "#ffffff";
            }
        );

        resetRootItem.addEventListener(
            "click",
            () => {

                if (!this.rootNodeId) {
                    return;
                }

                this.resetRoot();

                this.hideHierarchyContextMenu();
            }
        );

        menu.appendChild(resetRootItem);

        document.body.appendChild(menu);

        /*
        * Prevent the menu from going outside the screen.
        */
        const rect =
            menu.getBoundingClientRect();

        if (
            rect.right >
            window.innerWidth
        ) {
            menu.style.left =
                Math.max(
                    0,
                    window.innerWidth -
                    rect.width -
                    5
                ) + "px";
        }

        if (
            rect.bottom >
            window.innerHeight
        ) {
            menu.style.top =
                Math.max(
                    0,
                    window.innerHeight -
                    rect.height -
                    5
                ) + "px";
        }

        /*
        * Close menu when clicking somewhere else.
        */
        setTimeout(
            () => {

                document.addEventListener(
                    "click",
                    this.handleContextMenuOutsideClick,
                    {
                        once: true
                    }
                );

            },
            0
        );
    }

    private handleContextMenuOutsideClick =
        (event: MouseEvent): void => {

            if (
                this.contextMenu &&
                !this.contextMenu.contains(
                    event.target as Node
                )
            ) {
                this.hideHierarchyContextMenu();
            }
        };

    private hideHierarchyContextMenu(): void {

        if (this.contextMenu) {

            if (this.contextMenu.parentElement) {

                this.contextMenu.parentElement
                    .removeChild(
                        this.contextMenu
                    );
            }

            this.contextMenu = null;
        }
    };

    private drawLastRefresh(): void {

        /*
        * Chart title
        */
        if (this.chartTitle && this.chartTitle.trim() !== "") {

            const titleText =
                document.createElementNS(
                    "http://www.w3.org/2000/svg",
                    "text"
                );

            titleText.setAttribute("x", "15");
            titleText.setAttribute("y", "20");
            titleText.setAttribute(
                "font-family",
                "Segoe UI, Arial, sans-serif"
            );
            titleText.setAttribute("font-size", "16px");
            titleText.setAttribute("font-weight", "600");
            titleText.setAttribute("fill", "#333333");

            titleText.textContent = this.chartTitle;

            this.svg.appendChild(titleText);
        }


        /*
        * Last refresh date
        */
        if (!this.showLastUpdated || !this.lastRefresh) {
            return;
        }

        const refreshText =
            document.createElementNS(
                "http://www.w3.org/2000/svg",
                "text"
            );

        refreshText.setAttribute("x", "15");
        refreshText.setAttribute("y", "40");
        refreshText.setAttribute(
            "font-family",
            "Segoe UI, Arial, sans-serif"
        );
        refreshText.setAttribute("font-size", "12px");
        refreshText.setAttribute("fill", "#666666");

        refreshText.textContent =
            this.lastUpdatedLabel +
            " " +
            this.formatLastRefresh(this.lastRefresh);

        this.svg.appendChild(refreshText);
    }

    private formatLastRefresh(
        date: Date
    ): string {

        if (!date || isNaN(date.getTime())) {
            return "";
        }

        return new Intl.DateTimeFormat(
            "en-CA",
            {
                year: "numeric",
                month: "short",
                day: "numeric",
                hour: "numeric",
                minute: "2-digit",
                hour12: true
            }
        ).format(date);
    }
}