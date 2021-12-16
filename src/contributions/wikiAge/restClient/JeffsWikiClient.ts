import { RestClientBase } from "./RestClientBase";
import { IVssRestClientOptions } from "./Context";
import {VersionControlRecursionType} from "azure-devops-extension-api/Git";

export interface IGetWikiPageBatchRequest {
    continuationToken:string,
    top:number
}


export interface WikiPagesBatchResult {
    path:string;
    id:number;
}

export interface WikiPageVJSP {
    path:string;
    order:number;
    isParentPage:boolean;
    gitItemPath: string;
    subPages: any[];
    url: string;
    remoteUrl: string;
    id: number;
    content: any;
}

export class WikiPageBatchClient extends RestClientBase {
    constructor(options: IVssRestClientOptions) {
        super(options);
    }
    public static readonly RESOURCE_AREA_ID = "bf7d82a0-8aa5-4613-94ef-6172a5ea01f3";

    /**
     * Uploads an attachment on a comment on a wiki page.
     * 
     * @param content - Content to upload
     * @param project - Project ID or project name
     * @param wikiIdentifier - Wiki ID or wiki name.
     * @param pageId - Wiki page ID.
     */
    public async GetPagesBatch(        
        project: string,
        wikiIdentifier: string,
        continuationToken: string
        ): Promise<WikiPagesBatchResult[]> {

        let req:IGetWikiPageBatchRequest = {continuationToken:continuationToken,top:100};
        return this.beginRequest<WikiPagesBatchResult[]>({
            apiVersion: "5.2-preview.1",
            method: "POST",
            routeTemplate: "{project}/_apis/wiki/wikis/{wikiIdentifier}/pagesbatch",
            routeValues: {
                project: project,
                wikiIdentifier: wikiIdentifier
            },
            body: req
        });
    }


    public async getPageById(
        project: string,
        wikiIdentifier: string,
        id: number,
        recursionLevel?: VersionControlRecursionType,
        includeContent?: boolean
        ): Promise<WikiPageVJSP> {

        const queryValues: any = {
            recursionLevel: recursionLevel,
            includeContent: includeContent
        };

        return this.beginRequest<WikiPageVJSP>({
            apiVersion: "5.2-preview.1",
            httpResponseType: "application/json",
            routeTemplate: "{project}/_apis/wiki/wikis/{wikiIdentifier}/pages/{id}",
            routeValues: {
                project: project,
                wikiIdentifier: wikiIdentifier,
                id: id
            },
            queryParams: queryValues
        });
    }
}