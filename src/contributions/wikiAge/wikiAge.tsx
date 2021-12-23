import * as React from "react";

import * as SDK from "azure-devops-extension-sdk";
import * as API from "azure-devops-extension-api";
import { CommonServiceIds, IProjectPageService,IGlobalMessagesService, getClient, IProjectInfo } from "azure-devops-extension-api";

import { GitRestClient, GitItem } from "azure-devops-extension-api/Git";
import { WorkRestClient,BacklogConfiguration,BacklogLevelConfiguration, TeamFieldValues } from "azure-devops-extension-api/Work";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { WorkItemTypeReference } from "azure-devops-extension-api/WorkItemTracking/WorkItemTracking";
import { WorkItemTrackingRestClient, WorkItem } from "azure-devops-extension-api/WorkItemTracking";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

import { Page } from "azure-devops-ui/Page";
import { Table } from "azure-devops-ui/Table";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { Spinner, SpinnerSize } from "azure-devops-ui/Spinner";
import { Card } from "azure-devops-ui/Card";

import {Dropdown} from "azure-devops-ui/Dropdown";
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { IListBoxItem} from "azure-devops-ui/ListBox";
import { ZeroData } from "azure-devops-ui/ZeroData";
import { ListSelection } from "azure-devops-ui/List";

import {WikiRestClient, WikiV2} from "azure-devops-extension-api/Wiki";
import * as GetWiki from "./GetWiki"
import { showRootComponent } from "../../Common";
import {WikiPageBatchClient, WikiPagesBatchResult,WikiPageVJSP} from './restClient/JeffsWikiClient';

import * as GetGit from "./GitOps";
import * as TimeCalc from "./Time";
import * as GetProject from "./GetProjectInfo";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import {PageTableItem, IStatusIndicatorData, getStatusIndicatorData, CreateWorkItemButtonClick} from "./tableDataItems"
import { Status, Statuses, StatusSize } from "azure-devops-ui/Status";

import {        
    ITableColumn,
    SimpleTableCell,
    ColumnMore,
    ColumnSelect,
} from "azure-devops-ui/Table";
import { Link } from "azure-devops-ui/Link";
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { toggleFullScreen } from "azure-devops-ui/HeaderCommandBar";


interface IWikiAgeState {
    projectID:string;
    projectName:string;
    projectWikiID:string;
    projectWikiName:string;
    projectWikiRepoID:string;
    pageTableRows:PageTableItem[];
    doneLoading:boolean;
    daysThreshold:number;
    emptyWiki:boolean;
    renderOwners:boolean;    
    teamListItems:Array<IListBoxItem<{}>>;
    workItemType:WorkItemTypeReference|undefined;
    defaultDateSelection:DropdownSelection;
    defaultTeamSelection:DropdownSelection;
    selectedAreaPath:string;
}



export class WikiAgeContent extends React.Component<{}, IWikiAgeState> {
    private dateSelectionChoices = [                
        { text: "Updated in 30 Days", id: "30" },        
        { text: "Updated in 60 Days", id: "60" },        
        { text: "Updated in 90 Days", id: "90" },        
        { text: "Updated in 120 Days", id: "120" },
        { text: "Updated in 180 Days", id: "180" },        
        { text: "Updated in last 1 Year", id: "365" },        
        { text: "Updated in last 2 Years", id: "730" }
        
    ];


    public PageRef:WikiAgeContent;
    
    private wikiPageColumns:ITableColumn<PageTableItem>[]  = [
        {
            id: "statusCol",        
            readonly: true,
            renderCell: this.renderStatus,
            width: new ObservableValue(-6),
        },
        {        
            id: "workItemType",
            name: "Work Item",
            readonly: true,
            renderCell: this.RenderWorkItemButton,
            width: new ObservableValue(-28)
        },
        {
            id: "pagePath",
            name: "Page Path",
            readonly: true,
            renderCell: this.RenderIDLink,
            width: new ObservableValue(-125),
        },
        {
            id: "daysOld",
            name: "Days Since Update",
            readonly: true,
            renderCell: this.RenderDaysCount,
            width: new ObservableValue(-25),
        },    
        {
            id: "updateTimestamp",
            name: "Updated On",
            readonly: true,
            renderCell: this.RenderTimestamp,
            width: new ObservableValue(-45),
        },
        {
            id: "updatedBy",
            name: "Updated By",
            readonly: true,
            renderCell: this.RenderName,
            width: new ObservableValue(-35),
        },
    ]
    
    private wikiPageColumnsWitOwner : ITableColumn<PageTableItem>[] = [        
        {
            id: "statusCol",        
            readonly: true,
            renderCell: this.renderStatus,
            width: new ObservableValue(-6),
        },
     
        {        
            id: "workItemType",
            name: "Work Item",
            readonly: true,
            renderCell: this.RenderWorkItemButton,
            width: new ObservableValue(-28)
        },
        {
            id: "pagePath",
            name: "Page Path",
            readonly: true,
            renderCell: this.RenderIDLink,
            width: new ObservableValue(-120),
        },
        {
            id: "daysOld",
            name: "Days Since Update",
            readonly: true,
            renderCell: this.RenderDaysCount,
            width: new ObservableValue(-25),
        },    
        {
            id: "updateTimestamp",
            name: "Updated On",
            readonly: true,
            renderCell: this.RenderTimestamp,
            width: new ObservableValue(-45),
        },
        {
            id: "updatedBy",
            name: "Updated By",
            readonly: true,
            renderCell: this.RenderName,
            width: new ObservableValue(-30),
        },
        {
            id: "pageOwner",
            name: "Page Owner",
            readonly: true,
            renderCell: this.RenderOwner,
            width: new ObservableValue(-30),
        },
    ]    

    

    constructor(props:{}) {
        super(props);
        
        let initState:IWikiAgeState = this.initEmptyState();
        this.state = initState;
        this.PageRef = this;

    }

    public initEmptyState():IWikiAgeState
    {
        let defaultDateSelection = new DropdownSelection();
        defaultDateSelection.select(2);
    
        let initState:IWikiAgeState = {projectID:"", projectName:"", projectWikiID:"", projectWikiName:"", projectWikiRepoID:"", pageTableRows:[],doneLoading:false,daysThreshold:90, emptyWiki:true, renderOwners:false, teamListItems:[], workItemType:undefined, defaultDateSelection:defaultDateSelection, defaultTeamSelection:new DropdownSelection(), selectedAreaPath:""};
        return initState;
    }

        //After the Component has successfully loaded and is ready for processing, begin the initialization
    public async componentDidMount() {        
        await SDK.init();
        await SDK.ready();
        SDK.getConfiguration()
        SDK.getExtensionContext()
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        let newState:IWikiAgeState = this.initEmptyState();
        
        if(project) {
            newState.projectID = project.id;
            newState.projectName = project.name;
            
            try {
                let coreC:CoreRestClient = await this.GetCoreAPIClient();
                let workC:WorkRestClient = await this.GetWorkAPIClient();
                let teamList:WebApiTeam[] = await GetProject.getTeamsList(coreC,project);
                if(teamList.length > 0)
                {
                    let backlogConfig:BacklogConfiguration = await GetProject.GetTeamBacklogConfig(workC,teamList[0].id,project);
                    let workItemType:WorkItemTypeReference|undefined = this.GetRequirementWorkItemTypeForTeam(backlogConfig);
                    newState.workItemType = workItemType;                    
                }
                let listItems:Array<IListBoxItem<{}>> =this.GetTeamListItemsFromTeamList(teamList);
                newState.teamListItems = listItems;
                
                
                
                let teamIndexDefault:number = this.GetDefaultTeamIndex(listItems, project.name);
                
                let teamDefault:DropdownSelection = this.state.defaultTeamSelection;                
                teamDefault.select(teamIndexDefault);
                newState.defaultTeamSelection = teamDefault;
                let teamSetting:TeamFieldValues = await GetProject.GetTeamFieldValues(workC,project.id,newState.teamListItems[teamIndexDefault].id);
                newState.selectedAreaPath = teamSetting.defaultValue;
                this.setState(newState);
                await this.DoWork();                
            }
            catch(ex)
            {
                if(ex.message){
                    this.toastError(ex.message);
                    console.log(ex.message);
                }
                this.setState({doneLoading:true});
            }
        }
        else {
            this.toastError("Did not retrieve the project info");
        }
    }

    private GetTeamListItemsFromTeamList(teamList:WebApiTeam[]):Array<IListBoxItem<{}>>
    {
        let listItems:Array<IListBoxItem<{}>> = [];
        teamList.forEach((thisTeam) => {
            let t:IListBoxItem = {id:thisTeam.id, text:thisTeam.name};
            listItems.push(t);            
        });
        return listItems;
    }

    private GetRequirementWorkItemTypeForTeam(teamBacklogConfig:BacklogConfiguration):WorkItemTypeReference|undefined
    {
        let backlogTypeItem:WorkItemTypeReference|undefined = undefined;

        teamBacklogConfig.requirementBacklog.workItemTypes.forEach(thisReqType =>{

            if(thisReqType.name.toUpperCase() == "USER STORY" || thisReqType.name.toUpperCase() == "PRODUCT BACKLOG ITEM" || thisReqType.name.toUpperCase() == "REQUIREMENT")
            {
                backlogTypeItem = thisReqType;
            }
        });

        return backlogTypeItem;
    }

    private async DoWork(){
        let s:IWikiAgeState = this.state;

        try {
            let wclient = await this.GetWikiAPIClient();            
            let w:WikiV2 = await GetWiki.FindProjectWiki(wclient, s.projectID);
            let bclient = await this.GetBatchWikiAPIClient();            
            if(w)
            {   
                s.projectWikiID = w.id;                
                s.projectWikiName = w.name;
                s.projectWikiRepoID = w.repositoryId;
                
                let pgList:WikiPagesBatchResult[] = []
                try {
                    try 
                    {
                        pgList = await GetWiki.GetWikiPages(bclient,s.projectID,w.id);
                    }
                    catch(getEx)
                    {
                        
                        if(getEx)
                        {
                            if(JSON.stringify(getEx) != "{}")
                            {
                                this.toastError("Error Retrieving Wiki Page List: " +  JSON.stringify(getEx));
                            }
                        }

                    }
                    
                    let tblRow:PageTableItem[] = this.CollectPageRows(pgList,s.projectID,s.selectedAreaPath);

                    
                    let pageDetail:WikiPageVJSP[] = await GetWiki.GetPageDetails(bclient,s.projectID,w.id,pgList);                    
                    await this.MergePageDetails(tblRow, pageDetail);
                    tblRow.forEach(r => {
                        if(r.pageOwner.trim().length > 0)
                        {
                            s.renderOwners = true;
                        }                        
                    });
                    try {
                        let gitDetails:GitItem[] = await this.GetGitDetailsForPages(tblRow, w.repositoryId, s.projectID);
                        await this.MergeGitDetails(tblRow,gitDetails);
                    }
                    catch(ex)
                    {
                        this.toastError("git failed");
                    }

                    this.SetTableRowWorkItemType(tblRow,s.workItemType);
                    if(tblRow.length > 0)
                    {
                        s.emptyWiki=false;
                    }
                    tblRow = tblRow.sort(this.dateSort);
                    this.SetTableRowsDaysThreshold(tblRow, s.daysThreshold);
                    s.pageTableRows = tblRow;
                }
                catch(ex)
                {
                    this.toastError("During Work : " + ex + JSON.stringify(ex))
                }

                s.doneLoading=true;
                this.setState(s);
            }
        }
        catch(ex)
        {
            this.toastError(JSON.stringify(ex));
            this.setState({doneLoading:true});
        }

    }

    private async GetGitDetailsForPages(tableRows:PageTableItem[], projectWikiRepoID:string, projectID:string):Promise<GitItem[]>
    {

        let pagePaths:string[] = [];


        tableRows.forEach(t => {
            pagePaths.push(t.fileName);
        });        

        return new Promise<GitItem[]>(async (resolve,reject) => {  
            
            try{
                let gClient:GitRestClient = await this.GetGitAPIClient();
                let gitResults:GitItem[] = await GetGit.GetAllPageGitDetails(gClient,projectWikiRepoID, projectID,pagePaths);
                resolve(gitResults);
            }
            catch(ex)
            {
                this.toastError(JSON.stringify(ex));
                reject(ex);
            }

        });
        
        
    }

    
    private async SetTableRowWorkItemType(tableRows:PageTableItem[], wiType:WorkItemTypeReference|undefined)
    {

        let ndx:number = 0;

        if(tableRows.length > 0)
        {
            do {
                tableRows[ndx].workItemType = wiType;
                ndx++;
            } while(ndx < tableRows.length)
        }
    }

    private async MergePageDetails(tableRows:PageTableItem[], pageDetails:WikiPageVJSP[])
    {

        let ndx:number = 0;

        if(tableRows.length > 0)
        {
            do {
                let detail:WikiPageVJSP|undefined = this.GetPageDetailFromList(pageDetails,parseInt(tableRows[ndx].pageID));

                if(detail != undefined)
                {                    
                    tableRows[ndx].fileName = detail.gitItemPath;
                    tableRows[ndx].pageURL = detail.remoteUrl;                    
                    tableRows[ndx].pageOwner = this.FindOwnerTextInPage(detail.content);
                    tableRows[ndx].index = ndx;                    
                }    
                ndx++;
            } while(ndx < tableRows.length)
        }
    }

    private FindOwnerTextInPage(pageContent:string):string
    {
        let result:string = "";
        let ownerREx = /Owner:/gi;

        let pos = pageContent.search(ownerREx)
        if(pos == -1)
        {
            result ="" //no owner found
        }
        else
        {
           
           let contentLines:string[] = this.splitLines(pageContent);           
           let ownerFound:boolean = false;
           let ndx =0;
           if(contentLines.length >0)
           {
                do{
                    let ownerPos = contentLines[ndx].indexOf("Owner:");
                    if(ownerPos > -1)
                    {                     
                        result = contentLines[ndx].substring(ownerPos +6);                        
                        ownerFound = true;
                    }

                    ndx++;
                } while (!ownerFound && ndx < contentLines.length);
            }
        }
        

        return result.trim();
    }

    private splitLines(pageContent:string):string[]
    {
        let stringResult:string[] = [];

        stringResult = pageContent.split(/\r\n|\r|\n/);

        return stringResult
    }


    private SetTableRowsDaysThreshold(tableRows:PageTableItem[], threshold:number)
    {
        let ndx:number = 0;

        if(tableRows.length > 0)
        {
            do {
                tableRows[ndx].daysThreshold = threshold;
                ndx++;
            } while(ndx < tableRows.length)
        }
    }


    private async MergeGitDetails(tableRows:PageTableItem[], gitDetails:GitItem[])
    {
        let ndx:number = 0;

        let currentDate = new Date();
        if(tableRows.length > 0)
        {
            do {
                let detail:GitItem|undefined = this.GetGitItemFromList(gitDetails,tableRows[ndx].fileName);

                if(detail != undefined)
                {
                    
                    tableRows[ndx].updateTimestamp = detail.latestProcessedChange.committer.date.toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"long", day:"numeric"}) 
                    tableRows[ndx].updatedBy = detail.latestProcessedChange.committer.name;
                    tableRows[ndx].updateDateMili = detail.latestProcessedChange.committer.date.valueOf();
                    let timeDiffMili:number = currentDate.valueOf() - detail.latestProcessedChange.committer.date.valueOf();
                    let TimeDiffDuration:TimeCalc.IDuration = TimeCalc.getMillisecondsToTime(timeDiffMili);
                    tableRows[ndx].daysOld = TimeDiffDuration.days;

                }    
                ndx++;
            } while(ndx < tableRows.length)
        }
    }


    private GetGitItemFromList(allGitDetails:GitItem[], pathToFind:string):GitItem|undefined
    {
        let ndx:number = 0;
        let found:boolean = false;
        let result:GitItem|undefined = undefined;

        do{
            if(allGitDetails[ndx].path == pathToFind)
            {
                result = allGitDetails[ndx];
                found = true;
            }
            ndx++;
        } while (ndx < allGitDetails.length && !found)
        
        if(result == undefined)
        {
            this.toastError("Could not retrieve git info for path:" + pathToFind.toString());
        }

        return result;
    }
        
    private GetPageDetailFromList(allPages:WikiPageVJSP[], idToFind:number):WikiPageVJSP|undefined
    {
        let ndx:number = 0;
        let found:boolean = false;
        let result:WikiPageVJSP|undefined = undefined;
        do{
            if(allPages[ndx].id == idToFind)
            {
                result = allPages[ndx];
                found = true;
            }
            ndx++;
        } while (ndx < allPages.length && !found)
        
        if(result == undefined)
        {
            this.toastError("Could not retrieve details for Page ID:" + idToFind.toString());
        }

        return result;
    }


    ///Toast Error
    private async toastError(toastText:string)
    {
        const globalMessagesSvc = await SDK.getService<IGlobalMessagesService>(CommonServiceIds.GlobalMessagesService);
        globalMessagesSvc.addToast({        
            duration: 3000,
            message: toastText        
        });
        //this.setState({isToastVisible:true, isToastFadingOut:false, exception:toastText})
    }


    private async GetGitAPIClient():Promise<GitRestClient>
    {
        return new Promise<GitRestClient>(async (resolve,reject) => {
            try {
                let c:GitRestClient = API.getClient(GitRestClient ,);
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Git Client");
                reject(ex); 
            }
        });
    }

    private async GetWikiAPIClient():Promise<WikiRestClient>
    {
        return new Promise<WikiRestClient>(async (resolve,reject) => {
            try {
                let c:WikiRestClient = API.getClient(WikiRestClient);
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Wiki Client");
                reject(ex); 
            }
        });
    }

    private async GetBatchWikiAPIClient():Promise<WikiPageBatchClient>
    {
        return new Promise<WikiPageBatchClient>(async (resolve,reject) => {
            try {
                let c:WikiPageBatchClient = API.getClient(WikiPageBatchClient, {});
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Wiki Batch Client");
                reject(ex); 
            }
        });
    }

    private async GetCoreAPIClient():Promise<CoreRestClient>
    {
        return new Promise<CoreRestClient>(async (resolve,reject) => {
            try {
                let c:CoreRestClient = API.getClient(CoreRestClient, {});
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Wiki Batch Client");
                reject(ex); 
            }
        });
    }

    private async GetWorkAPIClient():Promise<WorkRestClient>
    {
        return new Promise<WorkRestClient>(async (resolve,reject) => {
            try {
                let c:WorkRestClient = API.getClient(WorkRestClient, {});
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Wiki Batch Client");
                reject(ex); 
            }
        });
    }

    public async GetWorkWITClient():Promise<WorkItemTrackingRestClient>
    {
        return new Promise<WorkItemTrackingRestClient>(async (resolve,reject) => {
            try {
                let c:WorkItemTrackingRestClient = API.getClient(WorkItemTrackingRestClient, {});
                resolve(c);
            }
            catch(ex)
            {
                this.toastError("Error while attempting to get Wiki Batch Client");
                reject(ex); 
            }
        });
    }


    
    private  selectTeam = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) =>{


        this.DoTeamSelect(item.id);

    }

    private async DoTeamSelect(teamId:string)
    {
        let projectId:string = this.state.projectID;
        let workClient:WorkRestClient = await this.GetWorkAPIClient();
        let teamSetting:TeamFieldValues = await GetProject.GetTeamFieldValues(workClient,projectId,teamId);
        let path = teamSetting.defaultValue;
        let pageTableRows:PageTableItem[] =this.state.pageTableRows;
        let ndx:number =0;
        do{

            pageTableRows[ndx].areaPath = path;
            ndx++;
        } while(ndx < pageTableRows.length)

        this.setState({pageTableRows:pageTableRows, selectedAreaPath:path});

    }

    //Day drop down select
    private SelectDays = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {

        let d:number = Number.parseInt(item.id);

        this.DoDateSelect(d);
    };

    private DoDateSelect(daysNumber:number)
    {
        let currentState:IWikiAgeState = this.state;
        currentState.daysThreshold = daysNumber;
        this.SetTableRowsDaysThreshold(currentState.pageTableRows, daysNumber);

        this.setState(currentState);

    }


    //////////////////////////////Table Render Items/////////////////////////////////////////


    private CollectPageRows(pageList:WikiPagesBatchResult[], projectId:string, areaPath:string):PageTableItem[]
    {
        let result:PageTableItem[] = [];
        pageList.forEach(thisPage => {
    
            let newPage:PageTableItem = {projectId:projectId, pageID:thisPage.id.toString(), pagePath:thisPage.path, fileName:thisPage.path, gitItemPath:"", pageURL:"", updateTimestamp: "", updatedBy:"", daysOld:-1, updateDateMili:-1, daysThreshold:90, pageOwner:"", workItemType:undefined, hasWorkItemCreated:false, index:0, pageRef:this, areaPath:areaPath, workItemNumber:"", workItemURL:""};        
            
            result.push(newPage);
        });
        return result;
    }
    
    private RenderIDLink(
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
    private renderStatus(
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
    
    
    private RenderDaysCount(
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
    private RenderTimestamp(
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
    
    private RenderName(
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
    
    private RenderOwner(
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
    
    private RenderWorkItemButton(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<PageTableItem>,
        tableItem: PageTableItem        
        ): JSX.Element
        {
        const {pagePath, daysOld, daysThreshold, pageOwner, updatedBy, workItemType,hasWorkItemCreated, index} = tableItem;
        
        if(workItemType)
        {
            let workItemTypetext:string =  workItemType.name;

            if(workItemTypetext.toUpperCase() == "PRODUCT BACKLOG ITEM")
            {
                workItemTypetext = "PBI";
            }

                     
            let btnText:string = "Create " + workItemTypetext;
            if( (daysOld >= daysThreshold)) 
            {                
                if(!hasWorkItemCreated)
                {                    
                    return (                    
                        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
                            <Button text={btnText} id={rowIndex.toString()} subtle={false} primary={true} onClick={() => {CreateWorkItemButtonClick(tableItem, rowIndex)} }></Button>
                        </SimpleTableCell>
                    );
                }
                else{
                    if(tableItem.workItemNumber == "" || tableItem.workItemURL == "")
                    {
                    return (                        
                        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
                            <Spinner label="Loading ..." size={SpinnerSize.xSmall} />
                        </SimpleTableCell>
                        );
                    }
                    else 
                    {
                        
                        return (
                        <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
                            <Link subtle={false} className="linkText"  excludeTabStop href={tableItem.workItemURL} target="_blank" tooltipProps={{ text: tableItem.workItemURL }}>
                                {workItemTypetext}: {tableItem.workItemNumber}
                            </Link>
                        </SimpleTableCell>
                        );
                    }
                }
            }
            else
            {
                return (
                    <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
                        
                    </SimpleTableCell>
                    );
            }
        }
        else{
    
            return (
            <SimpleTableCell columnIndex={columnIndex} tableColumn={tableColumn} key={"col-" + columnIndex} contentClassName="fontWeightSemiBold font-weight-semibold fontSizeM font-size-m scroll-hidden">
                
            </SimpleTableCell>
            );
        }
    }
    
    public dateSort(a:PageTableItem, b:PageTableItem) {
    
        if(a.updateDateMili < b.updateDateMili) {return -1;}
        if(a.updateDateMili > b.updateDateMili) {return 1;}
        return 0;
      }
    
    

    
    private getStatusIndicatorData(daysOld: number, threshhold:number): IStatusIndicatorData {
        
        
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
    
    






    public ResetRowState(rowIndex:number,data:PageTableItem)
    {

        let d = this.state.pageTableRows;
        d[rowIndex] = data;
        this.setState({pageTableRows:d});
    }




    private selection = new ListSelection({ selectOnFocus: false, multiSelect: true });

    //////////////////////////// End Table Render Items////////////////////////////////////////////


    private GetShowTeamBoxStyles(teamList:Array<IListBoxItem<{}>>):string
    {
        let suffix:string = "";
        if(teamList.length <= 1)
        {
            suffix = "-hidden";
        }
        return suffix;
    }

    private GetDefaultTeamIndex(teamList:Array<IListBoxItem<{}>>, projectName:string):number
    {
        let result:number = 0;
        let ndx:number = 0;        
        teamList.forEach(thisItem =>{
            
            if(thisItem.text)
            {                
                if(thisItem.text.indexOf(projectName) <= 1)
                {
                    result=ndx;
                }
            }
            ndx++;
        });        
        return result;
    }


    public render(): JSX.Element {

        let projectNameTitle:string = "Project: " + this.state.projectName;
        let wikiName = this.state.projectWikiName;
        let doneLoading:boolean = this.state.doneLoading;
        let renderOwners = this.state.renderOwners;
        let failedFindingWikiPages:boolean = this.state.emptyWiki;
        let tableColumns = [];
        let defaultDateSelection = this.state.defaultDateSelection;
        let teamList:Array<IListBoxItem<{}>> = this.state.teamListItems;
        let tableItemsNoIcons = new ArrayItemProvider<PageTableItem>(
            this.state.pageTableRows.map((item: PageTableItem) => {
                const newItem = Object.assign({}, item);
                //newItem.name = { text: newItem.name.text };
                return newItem;
            })
        );        
        let teamDefault:DropdownSelection = this.state.defaultTeamSelection;
        

        let classSuffix:string = this.GetShowTeamBoxStyles(teamList);
        let teamSelectPromptClass:string ="teamSelectPrompt" + classSuffix;
        let teamDropDownClass:string ="teamDropDown" + classSuffix;

        if(doneLoading)
        {
            if(failedFindingWikiPages)
            {
                return (<Page className="sample-hub flex-grow">
                        
                <Card className="selectionCard flex-row" >     
                <Header title="Wiki Age Report" titleSize={TitleSize.Large} />
                </Card>
                <ZeroData
                    primaryText="No Wiki Pages found in this project's Wiki"
                    secondaryText={
                        <span>
                           This report is designed to give you information about the wiki Pages in your project's wiki.  It will give you data when we are able to successfully find pages in this project's Wiki.
                        </span>
                    }
                    imageAltText="WikiAge"
                    imagePath={"./wikiAgeIcon.png"}
                    />
                
                </Page>
                );
            }
            else 
            {
                if(renderOwners)
                {
                    tableColumns = this.wikiPageColumnsWitOwner;
                }
                else
                {
                    tableColumns = this.wikiPageColumns;
                }
                return (
                    
                    <Page className="sample-hub flex-grow">
                        
                        <Card className="selectionCard flex-row" >     
                        <table style={{ textAlign:"left", minWidth:"650px", minHeight:""}}>
                            <tr>
                                <td style={{ minWidth:"200px"}}>
                                <Header title="Wiki Age Report" titleSize={TitleSize.Large} />
                                </td>
                                <td>
                                    <table>
                                        <tr>
                                            <td ><Header title="Age To Be Considered Old: " className="selectPrompt" titleSize={TitleSize.Small} /></td>
                                            <td><Dropdown items={this.dateSelectionChoices} placeholder="Select How old is old" ariaLabel="Basic" className="daysDropDown" onSelect={this.SelectDays} selection={defaultDateSelection} /> </td>
                                            <td>
                                                &nbsp;
                                            </td>
                                            <td>
                                                <Header title="Team To Add Work Items To: " className={teamSelectPromptClass} titleSize={TitleSize.Small} />
                                            </td>
                                            <td>
                                                <Dropdown items={teamList} placeholder="Select a Team" ariaLabel="Basic" className={teamDropDownClass} onSelect={this.selectTeam} selection={teamDefault} /> 
                                            </td>
                                        </tr>
                                    </table>
                                    
                                </td>
                                <td>

                                </td>
                            </tr>
                        </table>
                        </Card>
                        <Table
                            ariaLabel="Wiki Page Table"
                            columns={tableColumns}
                            itemProvider={tableItemsNoIcons}
                            role="table"
                            className="wiTable"
                            containerClassName="v-scroll-auto">

                        </Table>
                    </Page>
                );
            }
        }
        else {
            return(                      
                <Page className="sample-hub flex-grow">
                    <Header title="Wiki Age Report" titleSize={TitleSize.Large} />
                    <Header title={projectNameTitle} titleSize={TitleSize.Medium} />        
                    <Card className="flex-grow flex-center bolt-table-card" contentProps={{ contentPadding: true }}>       
                            <div className="flex-cell">
                                <div>
                                    <Spinner label="Loading ..." size={SpinnerSize.large} />
                                </div>
                            </div>          
                    </Card>
                </Page>
            );
        }
    }

}
showRootComponent(<WikiAgeContent />);