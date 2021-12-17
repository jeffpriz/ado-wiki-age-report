import * as React from "react";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { ISimpleListCell } from "azure-devops-ui/List";
import { MenuItemType } from "azure-devops-ui/Menu";
import { IStatusProps, Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import {WikiPagesBatchResult} from './restClient/JeffsWikiClient';
import {
    ColumnMore,
    ColumnSelect,
    ISimpleTableCell,
    renderSimpleCell,
    TableColumnLayout,
    ITableColumn,
    SimpleTableCell
} from "azure-devops-ui/Table";
import { Link } from "azure-devops-ui/Link";
import { css } from "azure-devops-ui/Util";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";


export const wikiPageColumns = [
    {
        id: "pagePath",
        name: "Page Path",
        readonly: true,
        renderCell: RenderIDLink,
        width: new ObservableValue(-85),
    },
    {
        id: "daysOld",
        name: "Days Since Update",
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(-20),
    },    
    {
        id: "updateTimestamp",
        name: "Updated On",
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(-55),
    },
    {
        id: "updatedBy",
        name: "Updated By",
        readonly: true,
        renderCell: renderSimpleCell,
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
}


export function CollectPageRows(pageList:WikiPagesBatchResult[]):PageTableItem[]
{

    let result:PageTableItem[] = [];
    pageList.forEach(thisPage => {

        let newPage:PageTableItem = {pageID:thisPage.id.toString(), pagePath:thisPage.path, fileName:"", gitItemPath:"", pageURL:"", updateTimestamp: "", updatedBy:"", daysOld:-1, updateDateMili:-1};        
        
        result.push(newPage);
    });

    console.log("TOTAL " + result.length.toString() + " page rows");
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
            <Link subtle={false} excludeTabStop href={pageURL} target="_blank" tooltipProps={{ text: pagePath }}>
                {pagePath}
            </Link>
        </SimpleTableCell>
    );
}

export function dateSort(a:PageTableItem, b:PageTableItem) {

    if(a.updateDateMili < b.updateDateMili) {return -1;}
    if(a.updateDateMili > b.updateDateMili) {return 1;}
    return 0;
  }