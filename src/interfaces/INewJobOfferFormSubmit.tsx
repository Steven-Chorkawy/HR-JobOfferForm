export interface INewJobOfferFormSubmit {
    JobID: string;
    Position: any;
    CandidateName: string;
    Department: any;
    JobType: string;        // ? This is a choice field so this might need to be a different type.

    // Array of objects that contains the template files that should be copied.
    TemplateFiles: any[];

    // Title = JobID-Position-CandidateName
    Title: string;          // * This field will be a combination of the other fields concatenated together.
}