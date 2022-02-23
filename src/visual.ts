/*
*  Power BI Visual CLI
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
"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import * as models from 'powerbi-models';
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import * as d3 from "d3";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import { VisualSettings } from "./settings";

export class Visual implements IVisual {
    private settings: VisualSettings;
    private host: IVisualHost;
    private d3visual: any;
    private currentFilterValues: Array<any>;
    private filterValuesWithDataTypes: Array<any>;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.d3visual = d3.select(options.element);
        this.d3visual.append('body');
        this.d3visual.selectAll('body').append('div').attr('class','searchbar-container')
        this.d3visual.selectAll('body').append('div').attr('class','scroll-container')
        this.currentFilterValues = new Array();
    }

    public update(options: VisualUpdateOptions) {
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.currentFilterValues = JSON.parse(this.settings.barslicer.slicerItems);
        this.d3visual.selectAll('.data-table').remove();
        this.d3visual.select('#erasorIcon').remove();
        if(this.settings.selectionControls.singleSelect){
            this.persistSelectAll(false);
        }
        if(this.settings.selectionControls.selectAll){
            this.persistSingleSelect(false);
        }
        var data = this.getDataset(options.dataViews[0]);
        this.persistTableName(data[0][3])
        this.persistColumnName(data[0][2])
        this.appendVisual(data);
        this.applySearchboxStyle();
    }

    public getDataset(dataView: DataView): any {
        var categoricalDataView = dataView.categorical;
        var categoricalValues = categoricalDataView.values[0].values;
        var categoricalValuesFormatString = categoricalDataView.values[0].source.format;
        var categoricalNames = categoricalDataView.categories[0].values;

        // Only if we have hierarchical structure with virtual table, take table name from identityExprs
        // Power BI creates hierarchy for date type of data (Year, Quater, Month, Days)
        // For it, Power BI creates a virtual table and gives it generated name, for example 'LocalDateTable_bcfa94c1-7c12-4317-9a5f-204f8a9724ca'
        // Visuals have to use a virtual table name as a target of JSON to filter date hierarchy properly

        if (categoricalDataView.categories[0].source.expr["ref"] === undefined) {
            var categoricalNamesColummName = categoricalDataView.categories[0].source.expr["level"];
            var categoricalNamesTableName = (<any>(<any>categoricalDataView.categories[0].source).identityExprs[(<any>categoricalDataView.categories[0].source).identityExprs.length - 1]).source.entity;
        }
        else {
            var categoricalNamesColummName = categoricalDataView.categories[0].source.expr["ref"];
            var categoricalNamesTableName = categoricalDataView.categories[0].source.expr["source"].entity;
        }

        var data = new Array();
        this.filterValuesWithDataTypes = new Array();

        for (var i = 0; i < categoricalNames.length; i++) {
            if (categoricalNames[i] != null) {
                data.push([categoricalNames[i], categoricalValues[i], categoricalNamesColummName.toString(), categoricalNamesTableName.toString(), categoricalValuesFormatString])
                this.filterValuesWithDataTypes.push(categoricalNames[i]);
            }
        }
        return data
    }

    public sortDataset(data: any): any {
        data.sort(function (a, b) { return d3.descending(a[1], b[1]) });
        return data;
    }

    public formatMeasure(measure: any, formatString: any): any {
        let iValueFormatter = valueFormatter.create({ format: formatString });
        return iValueFormatter.format(measure);
    }

    public appendVisual(data: any): any {
        var self = this;
        var maxValue = d3.max(data, function (d) { return Number(d[1]); })
        var minValue = d3.min(data, function (d) { return Number(d[1]); })

        // Setup the scale for the values for display, use abs max as max value
        var x = d3.scaleLinear()
            .domain([0, d3.max(data, function (d) { return Math.abs(d[1]); })])
            .range([0, 100]);

        // Erasor
        this.d3visual
            .selectAll('.searchbar-container')
            .append('a')
            .attr('class', 'erasor-icon')
            .attr('id', 'erasorIcon') as HTMLAnchorElement

        const erasor = document.getElementById('erasorIcon') as HTMLAnchorElement;
        erasor.addEventListener("click", this.onErasorClick.bind(this));

        // Search bar
        if(this.settings.selectionControls.searchBar && !this.d3visual.select('#searchInput').node()) {
            this.d3visual
                .selectAll('.searchbar-container')
                .append('a')
                .attr('class', 'search-icon')
                .attr('id', 'searchIcon')
                .style('font-size',this.settings.selectionControls.fontSize+"px")
                .style('font-family',this.settings.selectionControls.fontFamily)
                .style('color', this.settings.selectionControls.fontColor);

            this.d3visual
                .selectAll('.searchbar-container')
                .append('input')
                .attr('id', 'searchInput')
                .attr('class', 'searchbox')
                .attr('type', 'text')
                .attr('placeholder', 'Search')
                .attr('value', this.settings.barslicer.searchString)
                .style('font-size',this.settings.selectionControls.fontSize+"px")
                .style('font-family',this.settings.selectionControls.fontFamily)  
                .style('color', this.settings.selectionControls.fontColor);

            // Get the input box
            const input = document.getElementById('searchInput') as HTMLInputElement;

            let timeout = null;
            input.addEventListener('keyup', function (e) {
                clearTimeout(timeout);
                timeout = setTimeout(function () {
                    self.onSearchInput(input.value);
                }, 500);
            });
        }
        if(!this.settings.selectionControls.searchBar){
            this.d3visual.select('#searchInput').remove();
            this.d3visual.select('#searchIcon').remove();
        }

        // Table 
        var table = this.d3visual.selectAll('.scroll-container')
                    .append('table').attr('class','data-table')
                    .style('font-size',this.settings.selectionControls.fontSize+"px")
                    .style('font-family',this.settings.selectionControls.fontFamily);        

        // Select all row
        if (this.settings.selectionControls.selectAll) {
            var selectAllRow = table.append('tr').attr("class", "selectAllRow").attr("id", "selectAllRow");
            var span = selectAllRow.append('td').attr('class', 'checkbox').append('span').attr('class', 'checkmark')
            .style("height", Number(this.settings.selectionControls.fontSize)-1+"px")
            .style("width", Number(this.settings.selectionControls.fontSize)-1+"px")
            .style("border-color", this.settings.selectionControls.checkboxBorderColor);

            if (this.settings.barslicer.selectAllSelected) {
                span.node().classList.add("checked");
                span.style("background-color", this.settings.selectionControls.checkboxFillColor)
                span.style("border-color", this.settings.selectionControls.checkboxFillColor);
            }

            selectAllRow.append('td')
            .attr('class', 'data name')
            .text("Select all")
            .style('color', this.settings.selectionControls.fontColor);

            const trselectAllRow = document.getElementById("selectAllRow") as HTMLTableRowElement;
            trselectAllRow.addEventListener("click", this.onCheckboxChange.bind(this, trselectAllRow));
        }
        
        // Create a table with rows and bind a data row to each table row
        var tr = table.selectAll("tr.data")
            .data(data)
            .enter()
            .append("tr")
            .attr("class", "datarow")
            .attr("id", function (d) { return d[0] })
            .attr("value", function (d) { return d[0] })
            .attr("columnName", function (d) { return d[2] })
            .attr("tableName", function (d) { return d[3] })
            .attr("name", "slicerRow")
            .style("display", function (d) {
                                        if( JSON.parse(self.settings.barslicer.excludeItemsBySearch.toLowerCase()).includes(d[0].toLowerCase())){
                                            return "none"
                                        }
                                        else{
                                            return "table-row"
                                        }
                                    });

        var checkbox = tr.append("td").attr("class", "checkbox");

        // Create the name column
        tr.append("td").attr("class", "data name")
            .style("max-width", this.settings.selectionControls.maxTextColumnWidth.toString()+"px")
            .text(function (d) { return d[0] })
            .style('color', this.settings.selectionControls.fontColor);
            
        // Create the value column
        //  tr.append("td").attr("class", "data value")
        //      .text(function (d) { return self.formatMeasure(d[1], d[4]) })

        // Create a column at the beginning of the table for the chart
        var chart = tr.append("td").attr("class", "chart");
        chart.append("span").attr("class", "tooltip").text(function (d) { return self.formatMeasure(d[1], d[4]) });

        var minValueScaled = x(Math.abs(minValue));
        var maxValueScaled = maxValue < 0 ? 0 : x(Math.abs(maxValue));

        // Div structure of the chart
        chart.append("div").attr("class", "container").style("width", minValue > 0 ? "0%" : (minValueScaled / (minValueScaled + maxValueScaled)) * 100 + "%").append("div").attr("class", "negative");
        chart.append("div").attr("class", "container").style("width", maxValue > 0 ? (maxValueScaled / (minValueScaled + maxValueScaled)) * 100 + "%" : "0%").append("div").attr("class", "positive");

        // Negative div bar
        tr.select("div.negative")
            .style("background-color", this.settings.selectionControls.negativeColor)
            .style("width", function (d) { return d[1] > 0 ? "0%" : x(Math.abs(d[1])) + "%"; });

        // Positive div bar
        tr.select("div.positive")
            .style("background-color", this.settings.selectionControls.positiveColor)
            .style("width", function (d) { return d[1] > 0 ? x(d[1]) + "%" : "0%"; });

        // Checkboxes
        checkbox.append("span")
            .attr("class", "checkmark")
            .attr("id", function (d) { return d[0] })
            .attr("value", function (d) { return d[0] })
            .attr("columnName", function (d) { return d[2] })
            .attr("tableName", function (d) { return d[3] })
            .attr("name", "slicerButton")
            .attr("type", "checkbox")
            .style("height", Number(this.settings.selectionControls.fontSize)-1+"px")
            .style("width", Number(this.settings.selectionControls.fontSize)-1+"px")
            .style('border-color', this.settings.selectionControls.checkboxBorderColor);

        var slicerRows = document.getElementsByName("slicerRow")

        for (var i = 0; i < slicerRows.length; i++) {
            const tr = slicerRows[i] as HTMLTableRowElement

            if (this.settings.barslicer.slicerItems.indexOf(tr.getAttribute("value")) > -1) {
                const checkmark = tr.getElementsByClassName('checkmark')[0] as HTMLElement;
                checkmark.className += " checked"
                checkmark.style["background-color"] = this.settings.selectionControls.checkboxFillColor;
                checkmark.style["border-color"] = this.settings.selectionControls.checkboxFillColor;
            }
            tr.addEventListener("click", this.onCheckboxChange.bind(this, tr))
        }
        return table;
    }

    public applySearchboxStyle() {
        if(this.settings.selectionControls.searchBar) {
            this.d3visual
            .selectAll('#searchIcon')
            .style('font-size',this.settings.selectionControls.fontSize+"px");

            this.d3visual
            .selectAll('#searchInput')
            .style('font-size',this.settings.selectionControls.fontSize+"px")
            .style('font-family',this.settings.selectionControls.fontFamily);
        }
        else{
            this.d3visual.selectAll('.scroll-container')
            .style('height', '100vh')
        }
        this.d3visual
        .selectAll('#erasorIcon')
        .style('font-size',this.settings.selectionControls.fontSize+"px");
    }

    public onCheckboxChange(tr: HTMLTableRowElement) {
        if (tr.id === "selectAllRow") {
            if (!tr.getElementsByClassName('checkmark')[0].classList.contains("checked")) {
                this.selectAllFilterValues();
                this.persistSelectedItems();
                this.persistSelectAllSelected(true);
                this.applyFilter();
            }
            else {
                this.removeAllFilterValues();
                this.persistSelectedItems();
                this.persistSelectAllSelected(false);
                this.removeFilter();
            }
        }
        else if (!tr.getElementsByClassName('checkmark')[0].classList.contains("checked")) {
            if(this.settings.selectionControls.singleSelect){
                this.removeAllFilterValues();
            }
            this.incrementCurrentFilterValues(tr.getAttribute("value"));
            this.persistSelectedItems();
            this.applyFilter();
        }
        else {
            this.decrementCurrentFilterValues(tr.getAttribute("value"));
            this.persistSelectedItems();
            if (this.currentFilterValues.length > 0) {
                this.applyFilter();
            }
            else {
                this.removeFilter();
            }
        }
    }

    public onSearchInput(input: string) {
        this.persistSearchString(input);
        var excludeItems = []
        var datarows = document.getElementsByClassName('datarow') as HTMLCollectionOf<HTMLElement>;
        this.settings.barslicer.searchString = input;
        
        for (var i = 0; i < datarows.length; i++) {
            if(!datarows[i].getAttribute('value').toLowerCase().includes(input.toLowerCase())){
                excludeItems.push(datarows[i].getAttribute('value'));
            }
        }
        this.persistExcludeItemsBySearch(JSON.stringify(excludeItems));
    }

    public onErasorClick() {
        this.removeAllFilterValues();
        this.persistSelectedItems();
        this.persistSelectAllSelected(false);
        this.removeFilter();
    }

    public applySearchFilter(element: HTMLElement){
        element.style.display = 'none';
    }

    public applyFilter() {
        const basicFilter: models.IBasicFilter = {
            $schema: "http://powerbi.com/product/schema#basic",
            target: {
                table: this.settings.barslicer.tableName,
                column: this.settings.barslicer.columnName
            },
            operator: "In",
            //important that columns are in correct data type
            values: this.filterValuesWithDataTypes.filter(element => this.currentFilterValues.indexOf(element.toString()) !== -1),
            filterType: models.FilterType.Basic
        };

        this.host.applyJsonFilter(basicFilter, "general", "filter", powerbi.FilterAction.merge);
    }

    public removeFilter() {
        this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.merge);
    }

    public selectAllFilterValues() {
        this.currentFilterValues = [];
        var slicerRows = document.getElementsByName("slicerRow")
        for (var i = 0; i < slicerRows.length; i++) {
            this.currentFilterValues.push(slicerRows[i].getAttribute("value"));
        }
    }

    public removeAllFilterValues() {
        this.currentFilterValues = [];
    }

    public incrementCurrentFilterValues(filterValue: String) {
        this.currentFilterValues.push(filterValue)
    }

    public decrementCurrentFilterValues(filterValue: String) {
        this.currentFilterValues = this.currentFilterValues.filter(item => item !== filterValue)
    }

    public persistTableName(tableName: String) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "tableName": tableName
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistColumnName(columnName: String) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "columnName": columnName
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSelectedItems() {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "slicerItems": JSON.stringify(this.currentFilterValues)
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistExcludeItemsBySearch(excludeItemsBySearch: String) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "excludeItemsBySearch": excludeItemsBySearch
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSearchString(searchString: String) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "searchString": searchString
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSelectAllSelected(selected: boolean) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "selectAllSelected": selected
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSelectAll(selected: boolean) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "selectionControls",
                    selector: undefined,
                    properties: {
                        "selectAll": selected
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSingleSelect(selected: boolean) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "selectionControls",
                    selector: undefined,
                    properties: {
                        "singleSelect": selected
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    public persistSearchBarAlreadyCreated(selected: boolean) {
        let objects: powerbi.VisualObjectInstancesToPersist = {
            merge: [
                <VisualObjectInstance>{
                    objectName: "barslicer",
                    selector: undefined,
                    properties: {
                        "searchBarAlreadyCreated": selected
                    }
                }]
        };
        this.host.persistProperties(objects);
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        if (options.objectName === "barslicer") { return; }
        else {
            return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
        }
    }
}