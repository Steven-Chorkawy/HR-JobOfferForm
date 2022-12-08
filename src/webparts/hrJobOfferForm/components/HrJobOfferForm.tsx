import * as React from 'react';
import styles from './HrJobOfferForm.module.scss';
import { IHrJobOfferFormProps } from './IHrJobOfferFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FontSizes } from '@fluentui/theme/lib/fonts';
import { Field, FieldWrapper, Form, FormElement } from '@progress/kendo-react-form';
import { INewJobOfferFormSubmit } from '../../../interfaces/INewJobOfferFormSubmit';
import { TextField } from '@fluentui/react';
import { TaxonomyPicker } from '@pnp/spfx-controls-react';


export default class HrJobOfferForm extends React.Component<IHrJobOfferFormProps, {}> {


  //#region Form Fields
  private DepartmentInput = (fieldRenderProps: any) => {
    const { validationMessage, visited, label, id, valid, ...others } = fieldRenderProps;
    const showValidationMessage = visited && validationMessage;

    return (<FieldWrapper>
      <TaxonomyPicker
        allowMultipleSelections={false}
        termsetNameOrID={"Job Title"}
        label={label}
      />
    </FieldWrapper>);
  }
  //#endregion

  private _onSubmit = async (e: INewJobOfferFormSubmit): Promise<void> => {
    console.log('On Form Submit');
    console.log(e);
  }

  public render(): React.ReactElement<IHrJobOfferFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

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
                id={"Department"}
                name={"Department"}
                label={"Department"}
                component={this.DepartmentInput}
              />
            </FormElement>
          )}
        />
      </div>
    );
  }
}
