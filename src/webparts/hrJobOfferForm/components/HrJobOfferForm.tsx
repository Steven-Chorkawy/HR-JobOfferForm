import * as React from 'react';
import styles from './HrJobOfferForm.module.scss';
import { IHrJobOfferFormProps } from './IHrJobOfferFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FontSizes } from '@fluentui/theme/lib/fonts';
import { Field, FieldWrapper, Form, FormElement } from '@progress/kendo-react-form';
import { INewJobOfferFormSubmit } from '../../../interfaces/INewJobOfferFormSubmit';
import { DefaultButton, Dropdown, DropdownMenuItemType, IDropdownOption, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { FilePicker, IFilePickerResult, TaxonomyPicker } from '@pnp/spfx-controls-react';
import { MyTermSets } from '../../../enums/MyTermSets';
import { FormatTitle, GetTemplateDocuments } from '../../../HelperMethods/MyHelperMethods';


export default class HrJobOfferForm extends React.Component<IHrJobOfferFormProps, {}> {

  constructor(props: any) {
    super(props);
    GetTemplateDocuments();
  }

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
      />
    </FieldWrapper>);
  }

  private JobTypeDropDown = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, id, valid, ...others } = fieldRenderProps;
    const showValidationMessage = visited && validationMessage;
    // TODO: Replace this with a list pulled from metadata.
    const options: IDropdownOption[] = [
      { key: 'Extension Letter', text: 'Extension Letter' },
      { key: 'Fire', text: 'Fire' },
      { key: 'Full Time Non-Affiliated', text: 'Full Time Non-Affiliated' },
      { key: 'grape', text: 'Grape' },
      { key: 'broccoli', text: 'Broccoli' },
      { key: 'carrot', text: 'Carrot' },
      { key: 'lettuce', text: 'Lettuce' },
    ];

    return (<Dropdown
      placeholder="Select a Job Type"
      label={label}
      options={options}
    />);
  }

  private TemplateFilePicker = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, valid, ...others } = fieldRenderProps;

    return (
      // This FilePicker should only show results from the JobOfferTemplates library.
      <FilePicker
        buttonIcon="FileImage"
        label={label}
        buttonLabel={"Select Template File"}
        onSave={(filePickerResult: IFilePickerResult[]) => { console.log(filePickerResult); }}
        onChange={(filePickerResult: IFilePickerResult[]) => { console.log(filePickerResult) }}
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
    );
  }

  //#endregion

  private _onSubmit = async (e: INewJobOfferFormSubmit): Promise<void> => {
    console.log('On Form Submit');
    console.log(e);
    alert('Form submit... Check console...');
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
                required={true}
                component={TextField}
              />
              <Field
                id={"Position"}
                name={"Position"}
                label={"* Position"}
                termsetNameOrID={MyTermSets.JobTitles}
                panelTitle={"Select Position"}
                component={this.ManagedMetadataInput}
                requied={true}
                onChange={value => {
                  formRenderProps.onChange('Position', { value: value.length > 0 ? value[0] : null });
                }}
              />
              <Field
                id={"CandidateName"}
                name={"CandidateName"}
                label={"* Candidate Name"}
                required={true}
                component={TextField}
              />
              <Field
                id={"Department"}
                name={"Department"}
                label={"Department"}
                termsetNameOrID={MyTermSets.Departments}
                panelTitle={"Select Department"}
                component={this.ManagedMetadataInput}
                required={false}
                onChange={value => {
                  formRenderProps.onChange('Department', { value: value.length > 0 ? value[0] : null });
                }}
              />
              <Field
                id={"JobType"}
                name={"JobType"}
                label={"Job Type"}
                component={this.JobTypeDropDown}
              />
              <Field
                id={"TemplateFile"}
                name={"TemplateFile"}
                label={"Select Template File"}
                component={this.TemplateFilePicker}
              />

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
}
