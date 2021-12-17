import { GitRestClient, GitItem,VersionControlRecursionType  } from "azure-devops-extension-api/Git";

export async function GetPageGitDetails(gClient:GitRestClient, repoID:string, projectID:string, pagePath:string):Promise<GitItem>
{

    return new Promise<GitItem>(async (resolve,reject) => {   
        
        try {
            let gitDetailResult:GitItem = await gClient.getItem(repoID,pagePath,projectID,undefined,VersionControlRecursionType.None,true,true,false,undefined,false,false);
            resolve(gitDetailResult);
        }
        catch(ex){
            reject("Failed to get Git details for " + pagePath + " -- " + JSON.stringify(ex));
        }
        
    });


    
}

export async function GetAllPageGitDetails(gClient:GitRestClient, repoID:string, projectID:string, pageList:string[]):Promise<GitItem[]>
{
    let results:Promise<GitItem>[] = [];
    return new Promise<GitItem[]>(async (resolve,reject) => {  

        try {
            pageList.forEach(p => {
                results.push(GetPageGitDetails(gClient,repoID,projectID,p));
            });
            
            let allResults:GitItem[] = await Promise.all(results);

            resolve(allResults);
        }
        catch(ex)
        {
            reject("Git Error " + JSON.stringify(ex));
        }
    });
}