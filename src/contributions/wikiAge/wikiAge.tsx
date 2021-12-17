import * as React from "react";

import * as SDK from "azure-devops-extension-sdk";
import * as API from "azure-devops-extension-api";
import { CommonServiceIds, IProjectPageService,IGlobalMessagesService, getClient, IProjectInfo } from "azure-devops-extension-api";

import { GitRestClient, GitItem } from "azure-devops-extension-api/Git";

import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

import { Page } from "azure-devops-ui/Page";
import { Table } from "azure-devops-ui/Table";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { Spinner, SpinnerSize } from "azure-devops-ui/Spinner";
import { Card } from "azure-devops-ui/Card";

import {WikiRestClient, WikiV2} from "azure-devops-extension-api/Wiki";
import * as GetWiki from "./GetWiki"
import { showRootComponent } from "../../Common";
import {WikiPageBatchClient, WikiPagesBatchResult,WikiPageVJSP} from './restClient/JeffsWikiClient';
import * as TableSetup from "./tableDataItems";
import * as GetGit from "./GitOps";
import * as TimeCalc from "./Time";

interface IWikiAgeState {
    projectID:string;
    projectName:string;
    projectWikiID:string;
    projectWikiName:string;
    projectWikiRepoID:string;
    pageTableRows:TableSetup.PageTableItem[];
    doneLoading:boolean;
}

class WikiAgeContent extends React.Component<{}, IWikiAgeState> {

    constructor(props:{}) {
        super(props);
        
        let initState:IWikiAgeState = this.initEmptyState();
        
        this.state = initState;

    }

    public initEmptyState():IWikiAgeState
    {
        let initState:IWikiAgeState = {projectID:"", projectName:"", projectWikiID:"", projectWikiName:"", projectWikiRepoID:"", pageTableRows:[],doneLoading:false};
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
            newState.projectName = project.name
            this.setState(newState);
            await this.DoWork();                
        }
        else {
            this.toastError("Did not retrieve the project info");
        }
        
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
                console.log(w.id);
                s.projectWikiName = w.name;
                s.projectWikiRepoID = w.repositoryId;

                let pgList:WikiPagesBatchResult[] = await GetWiki.GetWikiPages(bclient,s.projectID,w.id);
                let tblRow:TableSetup.PageTableItem[] = TableSetup.CollectPageRows(pgList);

                
                let pageDetail:WikiPageVJSP[] = await GetWiki.GetPageDetails(bclient,s.projectID,w.id,pgList);
                
                await this.MergePageDetails(tblRow, pageDetail);
                
                try {
                    let gitDetails:GitItem[] = await this.GetGitDetailsForPages(tblRow, w.repositoryId, s.projectID);
                    await this.MergeGitDetails(tblRow,gitDetails);
                }
                catch(ex)
                {
                    this.toastError("git failed");
                }

                tblRow = tblRow.sort(TableSetup.dateSort);
                s.pageTableRows = tblRow;
                s.doneLoading=true;
                this.setState(s);
            }
        }
        catch(ex)
        {
            this.toastError(JSON.stringify(ex));
        }

    }

    private async GetGitDetailsForPages(tableRows:TableSetup.PageTableItem[], projectWikiRepoID:string, projectID:string):Promise<GitItem[]>
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


    private async MergePageDetails(tableRows:TableSetup.PageTableItem[], pageDetails:WikiPageVJSP[])
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
                }    
                ndx++;
            } while(ndx < tableRows.length)
        }

        

    }


    private async MergeGitDetails(tableRows:TableSetup.PageTableItem[], gitDetails:GitItem[])
    {
        let ndx:number = 0;

        let currentDate = new Date();
        if(tableRows.length > 0)
        {
            do {
                let detail:GitItem|undefined = this.GetGitItemFromList(gitDetails,tableRows[ndx].fileName);

                if(detail != undefined)
                {
                    
                    tableRows[ndx].updateTimestamp = detail.latestProcessedChange.committer.date.toString();
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


    public async GetGitAPIClient():Promise<GitRestClient>
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

    public async GetWikiAPIClient():Promise<WikiRestClient>
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

    public async GetBatchWikiAPIClient():Promise<WikiPageBatchClient>
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



    public render(): JSX.Element {

        let projectNameTitle = "Project: " + this.state.projectName;
        let wikiName = this.state.projectWikiName;
        let doneLoading:boolean = this.state.doneLoading;

        let tableItemsNoIcons = new ArrayItemProvider<TableSetup.PageTableItem>(
            this.state.pageTableRows.map((item: TableSetup.PageTableItem) => {
                const newItem = Object.assign({}, item);
                //newItem.name = { text: newItem.name.text };
                return newItem;
            })
        );

        if(doneLoading)
        {
            return (
                <Page className="sample-hub flex-grow">
                    <Header title="Wiki Age Report" titleSize={TitleSize.Large} />
                    <Header title={projectNameTitle} titleSize={TitleSize.Medium} />

                    <Table
                        ariaLabel="Work Item Table"
                        columns={TableSetup.wikiPageColumns}
                        itemProvider={tableItemsNoIcons}
                        role="table"
                        className="wiTable"
                        containerClassName="v-scroll-auto">

                    </Table>
                </Page>
            );
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