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
        Position: {
            // '__metadata': { 'type': 'SP.Taxonomy.TaxonomyFieldValue' }, // idk why I can't have this line.  Docs say I need it but it only works if I remove it.
            'Label': input.Position.name,
            'TermGuid': input.Position.key, // key prop not the termSet prop.
            'WssId': -1
        },
        Department1: {
            // '__metadata': { 'type': 'SP.Taxonomy.TaxonomyFieldValue' },
            'Label': input.Department.name,
            'TermGuid': input.Department.key, // key prop not the termSet prop.
            'WssId': -1
        }
    });

    // Copy template files. 
    if (input.TemplateFiles) {
        for (let templateFileIndex = 0; templateFileIndex < input.TemplateFiles.length; templateFileIndex++) {
            const templateFile = input.TemplateFiles[templateFileIndex];
            // use input.Title instead of templateFile.fileName.  The document should have the same name as the document set.
            // '- Offer of Employment' is a request from HR.
            let templateFileName = `${input.Title}${'- Offer of Employment'}${GetFileExtension(templateFile.fileName, templateFile.fileNameWithoutExtension)}`;
            await CopyTemplateDocument(input.Title, templateFile.fileAbsoluteUrl, templateFileName);
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

export const GetFileExtension = (fileNameWithExtension: string, fileNameWithoutExtension: string) => {
    // Produces an array of two elements.  The first should be empty.  The second should include the file extension.
    return fileNameWithExtension.split(fileNameWithoutExtension)[1];
}

/**
 * Get a list of Job Type
 */
export const GetJobTypes = async () => {
    let sp = getSP();
    let output = await sp.web.lists.getByTitle(MyLibraries.JobOffersLibrary).fields.getByInternalNameOrTitle('JobType').select('Choices')();
    return output["Choices"];
}

//#region Formatting Methods
// Title = JobID-Position-CandidateName
export const FormatTitle = (JobID: string, Position: string, CandidateName: string): string => {
    if (JobID === undefined || Position === undefined || CandidateName === undefined)
        return null;

    return `${JobID} - ${Position} - ${CandidateName}`;
}
export const FormatDocumentSetPath = (jobOfferTitle: string): string => `https://claringtonnet.sharepoint.com/sites/HR/JobOffers/${encodeURIComponent(jobOfferTitle)}`;
//#endregion

