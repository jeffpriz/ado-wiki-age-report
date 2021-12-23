import * as React from "react";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { IStatusProps, Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import { WorkItemTypeReference, WorkItem } from "azure-devops-extension-api/WorkItemTracking/WorkItemTracking";
import {WikiPagesBatchResult} from './restClient/JeffsWikiClient';
import {    
    ISimpleTableCell,    
    ITableColumn,
    SimpleTableCell
} from "azure-devops-ui/Table";
import { Link } from "azure-devops-ui/Link";
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import {WikiAgeContent} from './wikiAge';
import * as WI from './WorkItemFunctions';



export interface PageTableItem  {
    projectId:string;
    pageID: string;
    pagePath: string;
    fileName: string;
    gitItemPath:string;
    pageURL:string;
    updateTimestamp:string;
    updateDateMili:number;
    updatedBy:string;
    pageOwner:string;
    daysOld:number;
    daysThreshold:number;
    workItemType:WorkItemTypeReference|undefined;
    hasWorkItemCreated:Boolean;
    index:number;
    pageRef:WikiAgeContent;
    areaPath:string;
    workItemNumber:string;
    workItemURL:string;
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
            size={StatusSize.m}
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

export function RenderOwner(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<PageTableItem>,
    tableItem: PageTableItem
    ): JSX.Element
    {
    const { pageOwner} = tableItem;
    return (
    
        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
            <span className="tableText">{pageOwner}</span>
        </SimpleTableCell>
    );
}

//export function RenderWorkItemButton(
//    rowIndex: number,
//    columnIndex: number,
//    tableColumn: ITableColumn<PageTableItem>,
//    tableItem: PageTableItem
//    ): JSX.Element
//    {
//    const {pagePath, daysOld, daysThreshold, pageOwner, updatedBy, workItemType} = tableItem;
//    if(workItemType)
//    {
//        let wiType:string = workItemType.name;
//
//        let btnText:string = "Create " + wiType;
//        if( daysOld >= daysThreshold)
//        {
//            return (
//            
//                <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
//                    <ButtonGroup><Button text={btnText} id={rowIndex.toString()} subtle={false} primary={true} onClick={() => CreateWorkItemButtonClick(pagePath)}></Button></ButtonGroup>
//                </SimpleTableCell>
//            );
//        }
//        else
//        {
//            return (
//                <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
//                    
//                </SimpleTableCell>
//                );
//        }
//    }
//   else{
//
//        return (
//        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
//            
//        </SimpleTableCell>
//        );
//    }
//}

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

    if(daysOld >= threshhold )
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
    
export async function CreateWorkItemButtonClick(data:PageTableItem, rowIndex:number):Promise<any>
{


    data.hasWorkItemCreated= true;
    data.pageRef.ResetRowState(rowIndex,data);
    let WITCl = await data.pageRef.GetWorkWITClient();
    let wiName:string =  "";
    if(data.workItemType)
    {
        wiName = data.workItemType.name;        
    }
    else{
        wiName = "Task";
    }

    let workItemDescriptionText:string = GetWorkItemDescriptionText(data.pagePath, data.pageURL, data.updatedBy, data.updateTimestamp, data.pageOwner);
    let workItemTitle:string = `Update Wiki Document ${data.fileName}`;
    let acceptanceText:string = `<div>The Wiki document has been reviewed and updated to be current as of Today.</div> <div><br></div><div>Update the document with an Updated date or a Reviewed Date remark to make sure it comes off the Wiki Age Report`;
    try
    {
        let newWI:WorkItem = await WI.CreateWorkItem(WITCl,data.projectId,wiName,workItemTitle,workItemDescriptionText,acceptanceText, data.areaPath);
        data.hasWorkItemCreated = true;
        data.workItemNumber = newWI.id.toString();
        data.workItemURL = newWI._links.html.href;        
        data.pageRef.ResetRowState(rowIndex,data);
    }
    catch(ex)
    {
        data.hasWorkItemCreated = true;
        data.workItemNumber = "failed";
        data.workItemURL = "failed";
        console.log(ex.message);
    }
    //alert("New User Story Created");
    
}

export function GetWorkItemDescriptionText(documentPath:string, documentURL:string, lastUpdatedBy:string, lastUdpatedOn:string, documentOwner:string):string
{
    let documentownertext:string = "";
    if(documentOwner.trim().length >0)
    {
        documentownertext = `<div>The Document Owner is: ${documentOwner}</div>`;
    }
    let descriptionBody:string = `<div>As a team member, I would like to ensure that the Wiki Document is up to date</div><div><br></div><div><a href=\" ${documentURL}\"> ${documentPath}</a></div><div><br></div><div>Please review the documentation to see if there are any things in the document that need to be updated and make the necessary changes.</div><div><br></div>${documentownertext}<div>The Document was last updated on ${lastUdpatedOn} by ${lastUpdatedBy}</div>`;

    return descriptionBody;
}