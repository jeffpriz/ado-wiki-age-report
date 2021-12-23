import {WorkRestClient, BacklogConfiguration, TeamFieldValues, Board, BoardColumnType, BacklogLevelConfiguration} from "azure-devops-extension-api/Work";
import {CoreRestClient, WebApiTeam, TeamContext } from "azure-devops-extension-api/Core";
import {  IProjectInfo } from "azure-devops-extension-api";

    ///GetTeamConfig -- gets team configuration for Backlog and area
    export async function GetTeamBacklogConfig(workClient:WorkRestClient, teamID:string, project: IProjectInfo):Promise<BacklogConfiguration> 
    {

        return new Promise<BacklogConfiguration>(async (resolve,reject) => {  
            try 
            {
                if(project)
                {
                    let tc:TeamContext = {projectId: project.id, teamId:teamID, project:"",team:"" };
                    let teamBacklogConfigPromise:BacklogConfiguration = await workClient.getBacklogConfigurations(tc);
                    
                    resolve(teamBacklogConfigPromise);
                }
            }
            catch(ex){
                reject(ex);
            }
        });
    }

        ///getTeamsList -- gets the teams that are a part of the teamProject.
        export async function getTeamsList(coreClient:CoreRestClient, project: IProjectInfo):Promise<WebApiTeam[]>
        {
            
            return new Promise<WebApiTeam[]>(async (resolve,reject) => {  
                try {               
                
                let teamsList:WebApiTeam[] = await coreClient.getTeams(project.id,false,500);
                resolve(teamsList);
        
                }
                catch(ex)
                {
                    reject(ex);
                }
            });
        }


            //Takes the backlog configuration information and returns the names of the Work Item types that are a part of the chosen backlog (User Story, PBI, Bug, Feature, Epic.. )
        export function GetWorkItemTypesForBacklog(levelConfig:BacklogLevelConfiguration):string[]
        {
            let result:string[] = [];

            
                if(levelConfig)
                {
                    levelConfig.workItemTypes.forEach((thisType) => {
                        result.push(thisType.name);
                        
                    });
                }

            return result;
        }


        export async function GetTeamFieldValues(workClient:WorkRestClient, projectId:string, teamID:string):Promise<TeamFieldValues>
        {
            return new Promise<TeamFieldValues>(async (resolve,reject) => {  
                try{
                    let tc:TeamContext = {projectId: projectId, teamId:teamID, project:"",team:"" };
                    let teamSetting:TeamFieldValues =await workClient.getTeamFieldValues(tc);                    
                    resolve(teamSetting);
                }
                catch(ex)
                {
                    reject(ex);
                }
            });
        }
