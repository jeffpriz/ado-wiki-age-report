import * as React from "react";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { IStatusProps, Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import {WikiPagesBatchResult} from './restClient/JeffsWikiClient';
import {    
    ISimpleTableCell,    
    ITableColumn,
    SimpleTableCell
} from "azure-devops-ui/Table";
import { Link } from "azure-devops-ui/Link";

export const wikiPageColumns = [
    {
        id: "statusCol",        
        readonly: true,
        renderCell: renderStatus,
        width: new ObservableValue(-6),
    },
    {
        id: "pagePath",
        name: "Page Path",
        readonly: true,
        renderCell: RenderIDLink,
        width: new ObservableValue(-95),
    },
    {
        id: "daysOld",
        name: "Days Since Update",
        readonly: true,
        renderCell: RenderDaysCount,
        width: new ObservableValue(-25),
    },    
    {
        id: "updateTimestamp",
        name: "Updated On",
        readonly: true,
        renderCell: RenderTimestamp,
        width: new ObservableValue(-65),
    },
    {
        id: "updatedBy",
        name: "Updated By",
        readonly: true,
        renderCell: RenderName,
        width: new ObservableValue(-55),
    },
]

export interface PageTableItem extends ISimpleTableCell {
    pageID: string;
    pagePath: string;
    fileName: string;
    gitItemPath:string;
    pageURL:string;
    updateTimestamp:string;
    updateDateMili:number;
    updatedBy:string;
    daysOld:number;
    daysThreshold:number;
}


export function CollectPageRows(pageList:WikiPagesBatchResult[]):PageTableItem[]
{
    let result:PageTableItem[] = [];
    pageList.forEach(thisPage => {

        let newPage:PageTableItem = {pageID:thisPage.id.toString(), pagePath:thisPage.path, fileName:"", gitItemPath:"", pageURL:"", updateTimestamp: "", updatedBy:"", daysOld:-1, updateDateMili:-1, daysThreshold:90};        
        
        result.push(newPage);
    });
    return result;
}

export function RenderIDLink(
rowIndex: number,
columnIndex: number,
tableColumn: ITableColumn<PageTableItem>,
tableItem: PageTableItem
): JSX.Element
{
const { pagePath,  pageURL} = tableItem;
return (

    <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
        <Link subtle={false} className="linkText"  excludeTabStop href={pageURL} target="_blank" tooltipProps={{ text: pagePath }}>
            {pagePath}
        </Link>
    </SimpleTableCell>
);
}
export function renderStatus(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<PageTableItem>,
    tableItem: PageTableItem
    ): JSX.Element
    {
    const { daysOld, daysThreshold} = tableItem;
    return (
    
        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
        <Status
            {...getStatusIndicatorData(daysOld, daysThreshold).statusProps}
            className="icon-large-margin"
            size={StatusSize.s}
        />
        </SimpleTableCell>

    );
}


export function RenderDaysCount(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<PageTableItem>,
    tableItem: PageTableItem
    ): JSX.Element
    {
    const { daysOld} = tableItem;
    return (
    
        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden cell-row-center">
            <span className="days-column">{daysOld}</span>
        </SimpleTableCell>
    );
}
export function RenderTimestamp(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<PageTableItem>,
    tableItem: PageTableItem
    ): JSX.Element
    {
    const { updateTimestamp} = tableItem;
    return (
    
        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
            <span className="tableText">{updateTimestamp}</span>
        </SimpleTableCell>
    );
}

export function RenderName(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<PageTableItem>,
    tableItem: PageTableItem
    ): JSX.Element
    {
    const { updatedBy} = tableItem;
    return (
    
        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
            <span className="tableText">{updatedBy}</span>
        </SimpleTableCell>
    );
}

export function dateSort(a:PageTableItem, b:PageTableItem) {

    if(a.updateDateMili < b.updateDateMili) {return -1;}
    if(a.updateDateMili > b.updateDateMili) {return 1;}
    return 0;
  }


  export interface IStatusIndicatorData {
    statusProps: IStatusProps;
    label: string;
}

  export function getStatusIndicatorData(daysOld: number, threshhold:number): IStatusIndicatorData {
    
    
    const indicatorData: IStatusIndicatorData = {
        label: "Success",
        statusProps: { ...Statuses.Success, ariaLabel: "Success" },
    };

    if(daysOld > threshhold )
    {            
        indicatorData.statusProps = { ...Statuses.Failed, ariaLabel: "Failed" };
        indicatorData.label = "Failed";
    }
    else if (daysOld > threshhold - 7)
    {
        indicatorData.statusProps = { ...Statuses.Warning, ariaLabel: "Warning" };
        indicatorData.label = "Warning";
    }
    return indicatorData;
}