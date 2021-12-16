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
        id: "pageID",
        name: "Page ID",
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(-15),
    },
    {
        id: "pagePath",
        name: "Page Path",
        readonly: true,
        renderCell: RenderIDLink,
        width: new ObservableValue(-55),
    },
    {
        id: "fileName",
        name: "File",
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(-55),
    }
]

export interface PageTableItem extends ISimpleTableCell {
    pageID: string;
    pagePath: string;
    fileName: string;
    gitItemPath:string;
    pageURL:string;
}


export function CollectPageRows(pageList:WikiPagesBatchResult[]):PageTableItem[]
{

    let result:PageTableItem[] = [];
    pageList.forEach(thisPage => {

        let newPage:PageTableItem = {pageID:thisPage.id.toString(), pagePath:thisPage.path, fileName:"", gitItemPath:"", pageURL:""};        
        
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
            <Link subtle={false} excludeTabStop href={pageURL} target="_blank">
                {pagePath}
            </Link>
        </SimpleTableCell>
    );
}