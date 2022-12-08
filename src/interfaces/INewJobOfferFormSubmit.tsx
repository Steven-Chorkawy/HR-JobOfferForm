export interface INewJobOfferFormSubmit {
    JobID: string;
    Position: string;       // ? This might have to be a different type because it is a managed metadata field. 
    CandidateName: string;
    Department: string;     // ? This might have to be a different type because it is a managed metadata field. 
    JobType: string;        // ? This is a choice field so this might need to be a different type.
    
    // Title = JobID-Position-CandidateName
    Title: string;          // * This field will be a combination of the other fields concatenated together.
}