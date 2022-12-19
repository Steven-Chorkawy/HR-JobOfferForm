import { MyFormStatus } from "../../../enums/MyFormStatus";

export interface IHrJobOfferFormState {
    templateFiles: any[];
    jobTypes?: any[];
    formStatus: MyFormStatus;
    formStatusMessage?: string;
    jobOfferTitle?: string; // The title of a document set.
    jobOfferPath?: string; // A formatted link to a newly created Job Offer.
}
