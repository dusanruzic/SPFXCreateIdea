import { WebPartContext } from "@microsoft/sp-webpart-base";
import { EnvironmentType } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { IListCollection } from "./IList";
import { IListFieldCollection } from "./IListField";
import { IListItemCollection } from "./IListItem";



export class SharePointServiceManager {
    public context: WebPartContext;
    public environmentType: EnvironmentType;
    public ideaListID: string;
    public newListItemId: number

    public setup(context: WebPartContext, environmentType: EnvironmentType,  ideaListID: string): void {
        this.context = context;
        this.environmentType = environmentType;
      
        this.ideaListID = ideaListID;
    }

    public get(relativeEndpointUrl: string): Promise<any> {
        console.log(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`);
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`, SPHttpClient.configurations.v1)
        .then(
            response => {
                return response.json()
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }

    public getLists(): Promise<IListCollection> {
        return this.get('/_api/lists');
    }

    public getListItems(listId: string, selectedFields?: string[]) : Promise<IListItemCollection>{
        return this.get(`/_api/lists/getbyid('${listId}')/items?$select=*,Author/Name,Author/Title&$expand=Author/Id,AttachmentFiles`);
    }

    public getListItem(listId: string, itemId: number){
        return this.get(`/_api/lists/getbyid('${listId}')/items(${itemId})?$select=*,Author/Name,Author/Title,LinkToIdea/Title&$expand=Author/Id,LinkToIdea/Id,AttachmentFiles`);
    }

    public getListItemVersions(listId: string, itemId: number){
        //return this.get(`/_api/lists/getbyid('${listId}')/items(${itemId})/versions?$select=*,Author/Name,Author/Title,LinkToSpec/Title&$expand=Author/Id,LinkToSpec/Id,AttachmentFiles`);
        return this.get(`/_api/lists/getbyid('${listId}')/items(${itemId})/versions?$select=*,Author/Name,Author/Title,LinkToIdea/Title&$expand=Author/Id,LinkToIdea/Id,AttachmentFiles`);

    }

    public getListItemsFIltered(listId: string, filterString: string) : Promise<IListItemCollection>{
        console.log(`/_api/lists/getbyid('${listId}')/items?$filter=IdeaStatus eq '${filterString}'`);
        return this.get(`/_api/lists/getbyid('${listId}')/items?$select=*,Author/Name,Author/Title&$expand=Author/Id,AttachmentFiles&$filter=ElSpecStatus eq '${filterString}'`);
    }
    

    public getListFields(listId: string, showHiddenField: boolean = false): Promise<IListFieldCollection>{
        return this.get(`/_api/lists/getbyid('${listId}')/fields${!showHiddenField ? '?$filter=Hidden eq false' : ''}`);
        
    }

    
    public getUserByID(userID: string): Promise<any> {
        return this.get(`/_api/web/getuserbyid(${userID})`);
    }

    
    
    
    public getUsers(): Promise<any> {
        return this.get(`/_api/web/siteusers`);
    }

    public createIdea(name, desc, formula){

        const body = JSON.stringify({
            '__metadata': {
                'type': 'SP.Data.IdeaListItem'
            },
            'Title': name,
            'Comment1': desc,
            'IdeaFormula': formula
        })
        console.log(name);
        console.log(desc);
        console.log(formula);
        console.log(this.context.pageContext.web.absoluteUrl);
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getbyid('${this.ideaListID}')/items`, SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
            },
            body: body
        })
        .then(
            response => {
                return response.json()
            }
        )
        .catch(error => {
            return Promise.reject(error);
        });
    }

    public returnNumberOfFiles() {
        return (<HTMLInputElement>document.getElementById('txtAttachements')).files!.length;
    }

    public uploadPicture(num) {
        let files = (<HTMLInputElement>document.getElementById('txtAttachements')).files;
        //let file = files![0];
        console.log('broj slika:' + files!.length);
        for (let i =0; i< files!.length; i++) {
            console.log(`idemo ${num+1} sliku`)
            console.log('prikaz files tog');
            console.log(files![num]);
            if (files![num] != undefined || files![num] != null){

                return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/lists/getbyid('${this.ideaListID}')/items('${this.newListItemId}')/AttachmentFiles/add(FileName='${files![num].name}')`, SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': ''
                    },
                    body: files![num]
                })
                .then(
                    response => {
                        console.log('uspeo sa slikom!');
                        return response.json()
                    }
                )
                .catch(error => {
                    console.log('greska');
                    return Promise.reject(error);
                    
                });
            }
        }
        


    }
     
    

}

const SharePointService = new SharePointServiceManager();

export default SharePointService;  //singleton pattern