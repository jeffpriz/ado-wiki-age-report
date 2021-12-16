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
} from "azure-devops-ui/Table";
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
        renderCell: renderSimpleCell,
        width: new ObservableValue(-55),
    },
    {
        id: "fileName",
        name: "File Name",
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(-55),
    }
]

export interface ITableItem extends ISimpleTableCell {
    pageID: string;
    pagePath: string;
    fileName: string;
    
}


export function CollectPageRows(pageList:WikiPagesBatchResult[]):ITableItem[]
{

    let result:ITableItem[] = [];
    pageList.forEach(thisPage => {

        let newPage:ITableItem = {pageID:thisPage.id.toString(), pagePath:thisPage.path, fileName:""};        
        
        result.push(newPage);
    });

    console.log("TOTAL " + result.length.toString() + " page rows");
    return result;

}