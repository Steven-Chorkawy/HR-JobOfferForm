import * as React from 'react';
import styles from './HrJobOfferForm.module.scss';
import { IHrJobOfferFormProps } from './IHrJobOfferFormProps';
import { IHrJobOfferFormState } from './IHrJobOfferFormState';
import { escape } from '@microsoft/sp-lodash-subset';
import { FontSizes } from '@fluentui/theme/lib/fonts';
import { Field, FieldWrapper, Form, FormElement } from '@progress/kendo-react-form';
import { INewJobOfferFormSubmit } from '../../../interfaces/INewJobOfferFormSubmit';
import { DefaultButton, Dropdown, DropdownMenuItemType, IDropdownOption, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { FilePicker, IFilePickerResult, TaxonomyPicker } from '@pnp/spfx-controls-react';
import { MyTermSets } from '../../../enums/MyTermSets';
import { CreateDocumentSet, FormatTitle, GetJobTypes, GetTemplateDocuments } from '../../../HelperMethods/MyHelperMethods';
import { getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';



export default class HrJobOfferForm extends React.Component<IHrJobOfferFormProps, IHrJobOfferFormState> {

  constructor(props: any) {
    super(props);
    this.state = {
      templateFiles: []
    };

    this.SP = getSP(this.props.context);
  }

  private SP: SPFI = null;

  //#region Form Fields
  private ManagedMetadataInput = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, id, valid, termsetNameOrID, panelTitle, required, onChange, ...others } = fieldRenderProps;
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
      />
    </FieldWrapper>);
  }

  private JobTypeDropDown = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, id, valid, onChange, ...others } = fieldRenderProps;

    return (<FieldWrapper>
      <Dropdown
        placeholder="Select a Job Type"
        label={label}
        options={this.state.jobTypes}
        onChange={(event, item) => onChange(item)}
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
    const { validationMessage, visited, label, valid, onSave, ...others } = fieldRenderProps;

    return (
      // This FilePicker should only show results from the JobOfferTemplates library.
      <FieldWrapper>
        <FilePicker
          buttonIcon="FileImage"
          label={label}
          buttonLabel={"Select Template File"}
          onSave={(filePickerResult: IFilePickerResult[]) => onSave(filePickerResult)}
          onChange={(filePickerResult: IFilePickerResult[]) => {
            console.log('onChange')
            console.log(filePickerResult);
          }}
          context={this.props.context}
          defaultFolderAbsolutePath={"https://claringtonnet.sharepoint.com/sites/HR/JobOfferTemplates"}
          hideRecentTab={true}
          hideWebSearchTab={true}
          hideStockImages={true}
          hideOrganisationalAssetTab={true}
          hideOneDriveTab={true}
          hideLocalUploadTab={true}
          hideLocalMultipleUploadTab={true}
          hideLinkUploadTab={true}
        />
      </FieldWrapper>
    );
  }
  //#endregion

  //#region Validators
  // If value is good return empty string.  If value is bad return an error message.
  private positionValidator = (value: any): string => value ? "" : "Please select a Position."
  private jobIDValidator = (value: any): string => {
    console.log(value);
    if (!value)
      return "Please enter a Job ID.  Job IDs cannot contain the following characters.  \" * : < > ? / \\ |";
    return ['"', '*', ':', '<', '>', '?', '/', '\\', '|'].some(v => { return value.includes(v); }) ? "Job ID cannot contain the following characters.  \" * : < > ? / \\ |" : "";
  }
  private candidateNameValidator = (value: any): string => {
    if (!value)
      return "Please enter a Candidate Name.  Candidate Names cannot contain the following characters.  \" * : < > ? / \\ |";
    return ['"', '*', ':', '<', '>', '?', '/', '\\', '|'].some(v => { return value.includes(v); }) ? "Candidate Name cannot contain the following characters.  \" * : < > ? / \\ |" : "";
  }
  //#endregion

  private _onSubmit = async (e: INewJobOfferFormSubmit): Promise<void> => {
    console.log('On Form Submit');
    console.log(e);

    e.Title = FormatTitle(e.JobID, e.Position.name, e.CandidateName);
    let output = await CreateDocumentSet(e);
  }

  public render(): React.ReactElement<IHrJobOfferFormProps> {
    return (
      <div>
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
                component={this.TemplateFilePicker}
                onSave={(value: any) => {
                  formRenderProps.onChange('TemplateFiles', { value: value });
                  this.setState({ templateFiles: value });
                }}
              />

              <div>
                Files selected: {this.state.templateFiles.length}
                <ul>
                  {this.state.templateFiles.map((item: any) => (
                    <li><a href={item.fileAbsoluteUrl} target="_blank">{item.fileName}</a></li>
                  ))}
                </ul>
              </div>

              <div>
                Test Title: {
                  FormatTitle(
                    formRenderProps.valueGetter('JobID'),
                    formRenderProps.valueGetter('Position') && formRenderProps.valueGetter('Position').name,
                    formRenderProps.valueGetter('CandidateName'))
                }
              </div>
              <div className="k-form-buttons" style={{ marginTop: "20px" }}>
                <Stack horizontal tokens={{ childrenGap: 40 }}>
                  <PrimaryButton text="Submit" type="submit" />
                  <DefaultButton
                    text="Clear"
                    onClick={e => {
                      e.preventDefault();
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
    });
  }
}
