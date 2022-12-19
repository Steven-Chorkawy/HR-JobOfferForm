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
import { IItems } from "@pnp/sp/items";


// Title = JobID-Position-CandidateName
export const FormatTitle = (JobID: string, Position: string, CandidateName: string): string => {
    if (JobID === undefined || Position === undefined || CandidateName === undefined)
        return null;

    return `${JobID} - ${Position} - ${CandidateName}`;
}

export const GetTemplateDocuments = async (): Promise<IItems[]> => {
    const sp = getSP();
    const templateDocuments = await sp.web.lists.getByTitle(MyLibraries.JobOfferTemplatesLibrary).items.select("Id", "Title", "FileLeafRef", "File/Length").expand("File/Length")();
    return templateDocuments;
}

/**
 * Create a document set in the JobOffer library.
 * @param input INewJobOfferFormSubmit
 */
export const CreateDocumentSet = async (input: INewJobOfferFormSubmit): Promise<void> => {
    const sp = getSP();

    const library = await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    const folderPath = `${library.RootFolder.ServerRelativeUrl}/${input.Title}`
    const newFolderResult = await sp.web.folders.addUsingPath(folderPath);
    const newFolderProperties = await sp.web.getFolderByServerRelativePath(newFolderResult.data.ServerRelativeUrl).listItemAllFields();

    // Assign document set metadata. 
    // TODO: Add other properties here.
    await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).items.getById(newFolderProperties.ID).update({
        ContentTypeId: MyLibraries.JobOfferDocumentSetContentTypeID,
        JobID: input.JobID,
        CandidateName: input.CandidateName,
        JobType: input.JobType,

    });

    // Copy template files. 
    if (input.TemplateFiles) {
        for (let templateFileIndex = 0; templateFileIndex < input.TemplateFiles.length; templateFileIndex++) {
            const templateFile = input.TemplateFiles[templateFileIndex];
            await CopyTemplateDocument(input.Title, templateFile.fileAbsoluteUrl, templateFile.fileName);
        }
    }
}

/**
 * Copy the provided template documents into a given document set.
 * @param destinationUrl Path to the document set which the templates will be copied into. 
 * @param templatePaths A strings that containing the path to the template files that will be copied.
 */
export const CopyTemplateDocument = async (documentSetTitle: string, templatePath: string, templateFileName: string) => {
    const sp = getSP();
    const destinationUrl = `/sites/HR/JobOffers/${documentSetTitle}/${templateFileName}`;
    await sp.web.getFileByUrl(templatePath).copyTo(destinationUrl, false);
}

/**
 * Get a list of Job Type
 */
export const GetJobTypes = async () => {
    let sp = getSP();
    let output = await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).fields.getByInternalNameOrTitle('JobType').select('Choices')();
    return output["Choices"];
}

export const FormatDocumentSetPath = (jobOfferTitle: string): string => `https://claringtonnet.sharepoint.com/sites/HR/JobOffers/${encodeURIComponent(jobOfferTitle)}`;
