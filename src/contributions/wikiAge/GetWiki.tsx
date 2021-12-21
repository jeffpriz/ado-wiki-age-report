import {WikiRestClient, WikiV2} from "azure-devops-extension-api/Wiki";
import {WikiPageBatchClient, WikiPagesBatchResult, WikiPageVJSP} from './restClient/JeffsWikiClient';
import {VersionControlRecursionType} from "azure-devops-extension-api/Git";


export async function FindProjectWiki(wClient:WikiRestClient, projectID:string):Promise<WikiV2>
{
    return new Promise<WikiV2>(async (resolve,reject) => {        
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


export async function GetPageDetails(wClient:WikiPageBatchClient, projectID:string, wikiID:string, pageList:WikiPagesBatchResult[]):Promise<WikiPageVJSP[]>
{
    let pagePromises:Promise<WikiPageVJSP>[] = [];    

    
    return new Promise<WikiPageVJSP[]>(async (resolve,reject) => {    
        pageList.forEach(currentPage => {
            let p:Promise<WikiPageVJSP> = wClient.getPageById(projectID,wikiID,currentPage.id,VersionControlRecursionType.None,true);
            pagePromises.push(p);
        });  //end foreach
    
        try 
        {
            let pageDetails:WikiPageVJSP[] = await Promise.all(pagePromises);    
            resolve(pageDetails);
        }
        catch(ex)
        {
            reject("Error while retrieving all page Details " + JSON.stringify(ex));
        }
    });
}


export async function GetWikiPages(wClient:WikiPageBatchClient, projectID:string, wikiID:string):Promise<WikiPagesBatchResult[]>
{
    return new Promise<WikiPagesBatchResult[]>(async (resolve,reject) => {        
        let pagesList:WikiPagesBatchResult[] = [];
        let batchCount = 0;
        let continuationToken = 1;
        try {
            if(wClient)
            {
                do {
                    let thisBatchList:WikiPagesBatchResult[] = await wClient.GetPagesBatch(projectID,wikiID,continuationToken.toString());
                    batchCount = thisBatchList.length;                    
                    pagesList = pagesList.concat(thisBatchList);
                    let lastItem = pagesList[pagesList.length-1];
                    continuationToken =  lastItem.id;
                }
                while (batchCount == 100)               
                
                //pagesList = await wClient.GetPagesBatch(projectID,wikiID,"1");

                resolve(pagesList);
            }
        }
        catch(ex)
        {
            reject(ex);
        }
    });
}

