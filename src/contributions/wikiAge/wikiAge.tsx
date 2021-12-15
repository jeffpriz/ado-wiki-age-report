import * as React from "react";
import { showRootComponent } from "../../Common";
import * as SDK from "azure-devops-extension-sdk";
import * as API from "azure-devops-extension-api";
import { CommonServiceIds, IProjectPageService,IGlobalMessagesService, getClient, IProjectInfo } from "azure-devops-extension-api";
import { GitServiceIds, IVersionControlRepositoryService } from "azure-devops-extension-api/Git/GitServices";
import { Page } from "azure-devops-ui/Page";
import { GitRestClient, GitPullRequest, PullRequestStatus, GitPullRequestSearchCriteria, GitBranchStats } from "azure-devops-extension-api/Git";
import {WikiRestClient, WikiV2} from "azure-devops-extension-api/Wiki";


interface IWikiAgeState {
    projectID:string;
    projectName:string;
    projectWikiID:string;
    projectWikiName:string;
    projectWikiRepoID:string;
    
}

class WikiAgeContent extends React.Component<{}, IWikiAgeState> {

    constructor(props:{}) {
        super(props);
        
        let initState:IWikiAgeState = this.initEmptyState();
        
        this.state = initState;

    }

    public initEmptyState():IWikiAgeState
    {
        let initState:IWikiAgeState = {projectID:"", projectName:"", projectWikiID:"", projectWikiName:"", projectWikiRepoID:""};
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
        
        console.log("Do Worlo");
        try {
            let w:WikiV2 = await this.FindProjectWiki(s.projectID);

            if(w)
            {
                console.log("have a wiki:");
                s.projectWikiID = w.id;
                s.projectWikiName = w.name;
                s.projectWikiRepoID = w.repositoryId;

                this.setState(s);

            }
        }
        catch(ex)
        {
            this.toastError(JSON.stringify(ex));
        }

    }
        

    private async FindProjectWiki(projectID:string):Promise<WikiV2>
    {
        return new Promise<WikiV2>(async (resolve,reject) => {
            let wClient = await this.GetWikiAPIClient();
            let thisProjectWiki:WikiV2;
            let success:boolean = false;
            if(wClient)
            {
                let wikiList:WikiV2[] = await wClient.getAllWikis(projectID);
                if(wikiList)
                {
                    if(wikiList.length >0)
                    {
                        wikiList.forEach((w)=> {
                            if(w.projectId == projectID)
                            {
                                console.log("found a wiki");
                                thisProjectWiki = w;
                                success =true;
                                resolve(thisProjectWiki);  
                            }
                        });
                    }
                }
            }
            if(!success)
            { 
                reject("no project wiki found")
            }
        });
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
                let c:GitRestClient = API.getClient(GitRestClient);
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
                this.toastError("Error while attempting to get Git Client");
                reject(ex); 
            }
        });
    }



    public render(): JSX.Element {

        let projectName = this.state.projectName
        let wikiName = this.state.projectWikiName
        return (
            <Page className="sample-hub flex-grow">
                <div>Hello World  - {projectName}</div>             
                <div>
                    WikiName - {wikiName}
                    <br></br>
                    WikiId - {this.state.projectWikiID}
                    <br></br>
                    WikiRepoID - {this.state.projectWikiRepoID}
                </div> 
            </Page>
        );
    }

}
showRootComponent(<WikiAgeContent />);