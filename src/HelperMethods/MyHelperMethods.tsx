import { spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/security";
import { MyLibraries } from "../enums/MyLibraries";
import { getSP } from "../webparts/hrJobOfferForm/pnpjsConfig";
import { INewJobOfferFormSubmit } from "../interfaces/INewJobOfferFormSubmit";


// Title = JobID-Position-CandidateName
export const FormatTitle = (JobID: string, Position: string, CandidateName: string): string => {
    if (JobID === undefined || Position === undefined || CandidateName === undefined)
        return null;

    return `${JobID} - ${Position} - ${CandidateName}`;
}

// TODO: Get a list of documents in the template library.
export const GetTemplateDocuments = async () => {

    let sp = getSP();

    let test = await spfi(sp).site.getDocumentLibraries('https://claringtonnet.sharepoint.com/sites/HR');
    console.log('libraries:')
    console.log(test);

    let lists = await spfi(sp).web.lists()
    console.log('lists');
    console.log(lists);

    let templateDocuments = await spfi(sp).web.lists.getByTitle(MyLibraries.JobOfferTemplatesLibrary).items.select("Id", "Title", "FileLeafRef", "File/Length").expand("File/Length")();


    console.log('Template Documents');
    console.log(templateDocuments);

    return templateDocuments;
}

export const CreateDocumentSet = async (input: INewJobOfferFormSubmit) => {
    let sp = getSP();   

    let library = await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    let folderPath = `${library.RootFolder.ServerRelativeUrl}/${input.Title}`

    let newFolderResult = await sp.web.folders.addUsingPath(folderPath);
  
    let newFolderProperties = await sp.web.getFolderByServerRelativePath(newFolderResult.data.ServerRelativeUrl).listItemAllFields();

    await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).items.getById(newFolderProperties.ID).update({
        ContentTypeId: MyLibraries.JobOfferDocumentSetContentTypeID
    });
}

/**
 * Get a list of Job Type
 */
export const GetJobTypes = async () => {
    let sp = getSP();
    let output = await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).fields.getByInternalNameOrTitle('JobType').select('Choices')();
    return output["Choices"];
}