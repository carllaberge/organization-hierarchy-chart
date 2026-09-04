/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import powerbi from "powerbi-visuals-api";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Formatting settings for hierarchy nodes.
 */
class NodeCardSettings extends FormattingSettingsCard {

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background color",
        value: { value: "#FFFFFF" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Border color",
        value: { value: "#000000" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        value: { value: "#000000" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 12
    });

    boxShape = new formattingSettings.ItemDropdown({
        name: "boxShape",
        displayName: "Box shape",
        items: [
            {
                value: "rounded",
                displayName: "Rounded rectangle"
            },
            {
                value: "rectangle",
                displayName: "Rectangle"
            },
            {
                value: "ellipse",
                displayName: "Ellipse"
            }
        ],
        value: {
            value: "rounded",
            displayName: "Rounded rectangle"
        }
    });

    useLevelColors = new formattingSettings.ToggleSwitch({
        name: "useLevelColors",
        displayName: "Use level colors",
        value: true
    });

    level0Color = new formattingSettings.ColorPicker({
        name: "level0Color",
        displayName: "Level 0 color",
        value: { value: "#D9EAF7" }
    });

    level1Color = new formattingSettings.ColorPicker({
        name: "level1Color",
        displayName: "Level 1 color",
        value: { value: "#E8F3E8" }
    });

    level2Color = new formattingSettings.ColorPicker({
        name: "level2Color",
        displayName: "Level 2 color",
        value: { value: "#FFF2CC" }
    });

    level3Color = new formattingSettings.ColorPicker({
        name: "level3Color",
        displayName: "Level 3 color",
        value: { value: "#FCE4D6" }
    });

    level4Color = new formattingSettings.ColorPicker({
        name: "level4Color",
        displayName: "Level 4 color",
        value: { value: "#E4DFEC" }
    });

    nodeWidth = new formattingSettings.NumUpDown({
        name: "nodeWidth",
        displayName: "Node width",
        value: 220,
        options: {
            minValue: {
                value: 120,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 400,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    horizontalSpacing = new formattingSettings.NumUpDown({
        name: "horizontalSpacing",
        displayName: "Horizontal spacing",
        value: 40,
        options: {
            minValue: {
                value: 0,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 200,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    verticalSpacing = new formattingSettings.NumUpDown({
        name: "verticalSpacing",
        displayName: "Vertical spacing",
        value: 60,
        options: {
            minValue: {
                value: 20,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 200,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    connectionColor = new formattingSettings.ColorPicker({
        name: "connectionColor",
        displayName: "Connection line color",
        value: { value: "#888888" }
    });

    connectionWidth = new formattingSettings.NumUpDown({
        name: "connectionWidth",
        displayName: "Connection line width",
        value: 2,
        options: {
            minValue: {
                value: 1,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 10,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    connectionStyle = new formattingSettings.ItemDropdown({
        name: "connectionStyle",
        displayName: "Connection line style",
        items: [
            {
                value: "solid",
                displayName: "Solid"
            },
            {
                value: "dashed",
                displayName: "Dashed"
            },
            {
                value: "dotted",
                displayName: "Dotted"
            }
        ],
        value: {
            value: "solid",
            displayName: "Solid"
        }
    });

    name: string = "node";

    displayName: string = "Node";

    slices: Array<FormattingSettingsSlice> = [
        this.backgroundColor,
        this.borderColor,
        this.fontColor,
        this.fontSize,
        this.boxShape,
        this.useLevelColors,
        this.level0Color,
        this.level1Color,
        this.level2Color,
        this.level3Color,
        this.level4Color,   
        this.nodeWidth,
        this.horizontalSpacing,
        this.verticalSpacing,
        this.connectionColor,
        this.connectionWidth,
        this.connectionStyle
    ];
}

class HeaderCardSettings extends FormattingSettingsCard {

    chartTitle = new formattingSettings.TextInput({
        name: "chartTitle",
        displayName: "Chart title",
        placeholder: "Organization Chart",
        value: "Organization Chart"
    });

    showLastUpdated = new formattingSettings.ToggleSwitch({
        name: "showLastUpdated",
        displayName: "Show last updated",
        value: true
    });

    lastUpdatedLabel = new formattingSettings.TextInput({
        name: "lastUpdatedLabel",
        displayName: "Last updated label",
        placeholder: "Last updated:",
        value: "Last updated:"
    });

    name: string = "header";

    displayName: string = "Chart Header";

    slices: Array<FormattingSettingsSlice> = [
        this.chartTitle,
        this.showLastUpdated,
        this.lastUpdatedLabel
    ];
}

/**
 * Visual formatting settings model.
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {

    nodeCard = new NodeCardSettings();

    headerCard = new HeaderCardSettings();

    cards = [
        this.headerCard,
        this.nodeCard
    ];
}
