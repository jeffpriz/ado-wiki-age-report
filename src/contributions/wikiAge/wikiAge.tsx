import * as React from "react";
import { showRootComponent } from "../../Common";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IProjectPageService,IGlobalMessagesService, getClient, IProjectInfo } from "azure-devops-extension-api";
import { GitServiceIds, IVersionControlRepositoryService } from "azure-devops-extension-api/Git/GitServices";
import { Page } from "azure-devops-ui/Page";

interface IWikiAgeState {
    projectID:string;
    projectName:string;
}

class WikiAgeContent extends React.Component<{}, IWikiAgeState> {

    constructor(props:{}) {
        super(props);
        
        let initState:IWikiAgeState = this.initEmptyState();
        
        this.state = initState;

    }

    public initEmptyState():IWikiAgeState
    {
        let initState:IWikiAgeState = {projectID:"", projectName:""};
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
            }
            else {
                this.toastError("Did not retrieve the project info");
            }
            
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

    public render(): JSX.Element {

        let projectName = this.state.projectName
        return (
            <Page className="sample-hub flex-grow">
                <div>Hello World  - {projectName}</div>              
            </Page>
        );
    }

}
showRootComponent(<WikiAgeContent />);