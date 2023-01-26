import * as React from 'react';
import styles from './HrJobOfferForm.module.scss';
import { IHrJobOfferFormProps } from './IHrJobOfferFormProps';
import { IHrJobOfferFormState } from './IHrJobOfferFormState';
import { escape } from '@microsoft/sp-lodash-subset';
import { FontSizes } from '@fluentui/theme/lib/fonts';
import { Field, FieldWrapper, Form, FormElement } from '@progress/kendo-react-form';
import { INewJobOfferFormSubmit } from '../../../interfaces/INewJobOfferFormSubmit';
import { DefaultButton, Dropdown, DropdownMenuItemType, IDropdownOption, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Stack, TextField } from '@fluentui/react';
import { FilePicker, IFilePickerResult, TaxonomyPicker } from '@pnp/spfx-controls-react';
import { MyTermSets } from '../../../enums/MyTermSets';
import { CreateDocumentSet, FormatDocumentSetPath, FormatTitle, GetJobTypes, GetTemplateDocuments, GET_INVALID_CHARACTERS } from '../../../HelperMethods/MyHelperMethods';
import { getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { MyFormStatus } from '../../../enums/MyFormStatus';
import { Card, CardActions, CardBody, CardSubtitle, CardTitle } from '@progress/kendo-react-layout';


export default class HrJobOfferForm extends React.Component<IHrJobOfferFormProps, IHrJobOfferFormState> {

  constructor(props: any) {
    super(props);
    this.state = {
      templateFiles: [],
      formStatus: MyFormStatus.New
    };

    this.SP = getSP(this.props.context);
  }

  private SP: SPFI = null;

  //#region Form Fields
  private ManagedMetadataInput = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, termsetNameOrID, panelTitle, required, onChange, ...others } = fieldRenderProps;
    const showValidationMessage = visited && validationMessage;

    return (<FieldWrapper>
      <TaxonomyPicker
        allowMultipleSelections={false}
        termsetNameOrID={termsetNameOrID}
        label={label}
        panelTitle={panelTitle}
        context={this.props.context}
        required={required}
        onChange={onChange}
        errorMessage={showValidationMessage}
        {...others}
      />
    </FieldWrapper>);
  }

  private JobTypeDropDown = (fieldRenderProps: any) => {
    const { label, onChange, ...others } = fieldRenderProps;

    return (<FieldWrapper>
      <Dropdown
        placeholder="Select a Job Type"
        label={label}
        options={this.state.jobTypes}
        onChange={(event, item) => onChange(item)}
        {...others}
      />
    </FieldWrapper>);
  }

  private TextFieldInput = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, id, name, ...others } = fieldRenderProps;
    const showValidationMessage = visited && validationMessage;
    return <FieldWrapper>
      <TextField
        id={id}
        name={name}
        label={label}
        errorMessage={showValidationMessage && validationMessage}
        {...others}
      />
    </FieldWrapper>;
  }

  private TemplateFilePicker = (fieldRenderProps: any) => {
    const { label, onSave, jobType, ...others } = fieldRenderProps;

    let folderPath = 'https://claringtonnet.sharepoint.com/sites/HR/JobOfferTemplates';
    if (jobType != undefined)
      folderPath = `${folderPath}/${jobType}`;

    return (
      // This FilePicker should only show results from the JobOfferTemplates library.
      <FieldWrapper>
        <FilePicker
          buttonIcon="FileImage"
          label={label}
          buttonLabel={"Select Template File"}
          onSave={(filePickerResult: IFilePickerResult[]) => onSave(filePickerResult)}
          context={this.props.context}
          defaultFolderAbsolutePath={folderPath}
          hideRecentTab={true}
          hideWebSearchTab={true}
          hideStockImages={true}
          hideOrganisationalAssetTab={true}
          hideOneDriveTab={true}
          hideLocalUploadTab={true}
          hideLocalMultipleUploadTab={true}
          hideLinkUploadTab={true}
          {...others}
        />
      </FieldWrapper>
    );
  }
  //#endregion

  //#region Validators
  // If value is good return empty string.  If value is bad return an error message.
  private positionValidator = (value: any): string => value ? "" : "Please select a Position."
  private jobIDValidator = (value: any): string => {
    if (!value)
      return "Please enter a Job ID.  Job IDs cannot contain the following characters.  \" * : < > ? / \\ | #";
    return GET_INVALID_CHARACTERS.some(v => { return value.includes(v); }) ? "Job ID cannot contain the following characters.  \" * : < > ? / \\ | #" : "";
  }
  private candidateNameValidator = (value: any): string => {
    if (!value)
      return "Please enter a Candidate Name.  Candidate Names cannot contain the following characters.  \" * : < > ? / \\ | #";
    return GET_INVALID_CHARACTERS.some(v => { return value.includes(v); }) ? "Candidate Name cannot contain the following characters.  \" * : < > ? / \\ | #" : "";
  }
  //#endregion

  private _onSubmit = async (e: INewJobOfferFormSubmit): Promise<void> => {
    try {
      e.Title = FormatTitle(e.JobID, e.Position.name, e.CandidateName);

      // Save the title in state to be used in success/failed error messages. 
      this.setState({
        jobOfferTitle: e.Title,
        jobOfferPath: FormatDocumentSetPath(e.Title),
        formStatus: MyFormStatus.Loading
      });

      await CreateDocumentSet(e)
        .then(value => {
          this.setState({
            formStatus: MyFormStatus.Success,
            formStatusMessage: "Job Offer has successfully been created!"
          });
        })
        .catch(reason => {
          console.log('Form failed to submit.');
          console.log(reason);
          this.setState({
            formStatus: MyFormStatus.Failed,
            formStatusMessage: "Something went wrong.  Failed to submit form."
          });
        });
    }
    catch (error) {
      console.log("Form failed to submit");
      console.log(error);
      this.setState({
        formStatus: MyFormStatus.Failed,
        formStatusMessage: "Something went wrong.  Failed to submit form."
      });
    }
  }

  public render(): React.ReactElement<IHrJobOfferFormProps> {
    return (
      <div style={{ paddingBottom: '2em', paddingLeft: '2em', paddingRight: '2em' }}>
        <div style={{ fontSize: FontSizes.size32 }}>{this.props.description}</div>
        <hr />
        <Form
          onSubmit={(e: INewJobOfferFormSubmit) => this._onSubmit(e)}
          render={(formRenderProps) => (
            <FormElement style={{ maxWidth: 650 }}>
              <Field
                id={"JobID"}
                name={"JobID"}
                label={"* Job ID"}
                component={this.TextFieldInput}
                validator={this.jobIDValidator}
              />
              <Field
                id={"Position"}
                name={"Position"}
                label={"* Position"}
                termsetNameOrID={MyTermSets.JobTitles}
                panelTitle={"Select Position"}
                component={this.ManagedMetadataInput}
                required={true}
                validator={this.positionValidator}
                onChange={value => formRenderProps.onChange('Position', { value: value.length > 0 ? value[0] : null })}
              />
              <Field
                id={"CandidateName"}
                name={"CandidateName"}
                label={"* Candidate Name"}
                component={this.TextFieldInput}
                validator={this.candidateNameValidator}
              />
              <Field
                id={"Department"}
                name={"Department"}
                label={"Department"}
                termsetNameOrID={MyTermSets.Departments}
                panelTitle={"Select Department"}
                component={this.ManagedMetadataInput}
                required={false}
                onChange={value => formRenderProps.onChange('Department', { value: value.length > 0 ? value[0] : null })}
              />
              <Field
                id={"JobType"}
                name={"JobType"}
                label={"Job Type"}
                component={this.JobTypeDropDown}
                onChange={value => formRenderProps.onChange('JobType', { value: value.text })}
              />
              <Field
                id={"TemplateFiles"}
                name={"TemplateFiles"}
                label={"Select Template Files"}
                jobType={formRenderProps.valueGetter('JobType')}
                component={this.TemplateFilePicker}
                onSave={(value: any) => {
                  formRenderProps.onChange('TemplateFiles', { value: value });
                  this.setState({ templateFiles: value });
                }}
              />

              <div>
                Files selected: {this.state.templateFiles.length}
                <ul>
                  {this.state.templateFiles.map(item => (
                    <li key={item.fileName}><a href={item.fileAbsoluteUrl} target="_blank" rel="noreferrer">{item.fileName}</a></li>
                  ))}
                </ul>
              </div>

              <div>
                <Card>
                  <CardBody>
                    <CardTitle>
                      {
                        FormatTitle(
                          formRenderProps.valueGetter('JobID'),
                          formRenderProps.valueGetter('Position') && formRenderProps.valueGetter('Position').name,
                          formRenderProps.valueGetter('CandidateName'))
                      }
                    </CardTitle>
                    {
                      (FormatTitle(
                        formRenderProps.valueGetter('JobID'),
                        formRenderProps.valueGetter('Position') && formRenderProps.valueGetter('Position').name,
                        formRenderProps.valueGetter('CandidateName')) === null) &&
                      <CardSubtitle>Please enter Job ID, Position, and Candidate Name.</CardSubtitle>
                    }
                    <div>
                      {
                        this.state.formStatus == MyFormStatus.Failed &&
                        <MessageBar messageBarType={MessageBarType.error}>{this.state.formStatusMessage}</MessageBar>
                      }
                      {
                        this.state.formStatus == MyFormStatus.Success &&
                        <MessageBar messageBarType={MessageBarType.success} isMultiline={true}>
                          <div><a href={this.state.jobOfferPath} target='_blank'>{this.state.jobOfferTitle}</a> has successfully been created!  Click the link to view Job Offer.</div>
                          <div style={{ paddingTop: '10px' }}><a href="https://claringtonnet.sharepoint.com/sites/HR/JobOffers" target='_blank'>Click here to view all job offers.</a></div>
                        </MessageBar>
                      }
                      {
                        this.state.formStatus == MyFormStatus.Loading &&
                        <ProgressIndicator label={`Creating ${this.state.jobOfferTitle}`} description="Creating document set, applying metadata, copying tempalte documents..." />
                      }
                    </div>
                  </CardBody>
                </Card>
              </div>



              <div className="k-form-buttons" style={{ marginTop: "20px" }}>
                <Stack horizontal tokens={{ childrenGap: 40 }}>
                  <PrimaryButton text="Submit" type="submit" />
                  <DefaultButton
                    text="Clear"
                    onClick={e => {
                      e.preventDefault();
                      this.setState({
                        formStatus: MyFormStatus.New,
                        formStatusMessage: null,
                        jobOfferTitle: null,
                        jobOfferPath: null,
                        templateFiles: []
                      });
                      formRenderProps.onFormReset();
                    }}
                  />
                </Stack>
              </div>
            </FormElement>
          )}
        />
      </div >
    );
  }

  public componentDidMount(): void {
    GetJobTypes().then(values => {
      this.setState({
        // mapped to be formatted for dropdowns.
        jobTypes: values.map(v => { return { key: v, text: v } })
      });
    }).catch(reason => {
      console.error(reason);
      alert('Failed to load Job Types.');
    });
  }
}
