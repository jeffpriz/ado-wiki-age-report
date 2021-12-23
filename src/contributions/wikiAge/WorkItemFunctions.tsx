import { IWorkItemFormNavigationService, WorkItemTrackingRestClient, WorkItemTrackingServiceIds, WorkItem } from "azure-devops-extension-api/WorkItemTracking";
import { JsonPatchDocument,JsonPatchOperation, Operation } from "azure-devops-extension-api/WebApi";

export const TITLE_FIELD:string = "System.Title";
export const DESCRIPTION_FIELD:string = "System.Description";
export const AREA_PATH_FIELD:string = "System.AreaPath";
export const ACCEPTANCE_CRITERIA_FIELD:string = "Microsoft.VSTS.Common.AcceptanceCriteria";
export const WORK_ITEM_TYPE_FIELD:string ="System.WorkItemType";
export const STACK_RANK_FIELD:string = "Microsoft.VSTS.Common.StackRank";
export const AREA_ID_FIELD:string = "System.AreaId";
export const ASSIGNED_TO_FIELD:string = "System.AssignedTo";


export async function CreateWorkItem(witClient:WorkItemTrackingRestClient, projectId:string, workItemType:string, title:string, description:string, criteria:string, areaPath:string):Promise<WorkItem>
{

    let patch:JsonPatchDocument = []
    let ops:JsonPatchOperation[] = [];
    
    var wiPatchNameOp: JsonPatchOperation = {op: Operation.Add, path:"/fields/" + TITLE_FIELD, value:title, from:""};
    var wiPatchDescOp: JsonPatchOperation = {op: Operation.Add, path:"/fields/" + DESCRIPTION_FIELD, value:description,from:""};
    var wiPatchAreaPathOp: JsonPatchOperation = {op: Operation.Add, path:"/fields/" + AREA_PATH_FIELD, value:areaPath,from:""};
    var wiPatchAcceptOp: JsonPatchOperation = {op: Operation.Add, path:"/fields/" + ACCEPTANCE_CRITERIA_FIELD, value:criteria, from:""};

    ops.push(wiPatchNameOp);
    ops.push(wiPatchDescOp);
    ops.push(wiPatchAreaPathOp);
    ops.push(wiPatchAcceptOp);
    patch = ops;

    
    return new Promise<WorkItem>(async (resolve, reject) => {
        try
        { 
            let r:WorkItem = await witClient.createWorkItem(patch,projectId, workItemType);
            resolve(r);
        }
        catch(ex)
        {
            reject(ex);
        }
    });
}

