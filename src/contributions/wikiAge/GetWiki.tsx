import {WikiRestClient, WikiV2} from "azure-devops-extension-api/Wiki";
import {WikiPageBatchClient, WikiPagesBatchResult} from './restClient/JeffsWikiClient';
import * as SDK from "azure-devops-extension-sdk";
import * as API from "azure-devops-extension-api";
import { CommonServiceIds, IProjectPageService,IGlobalMessagesService, getClient, IProjectInfo } from "azure-devops-extension-api";


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
            reject("Failed getting the pages list" +  JSON.stringify(ex));
        }
    });
}

