import * as React from 'react';
import styles from './HrJobOfferForm.module.scss';
import { IHrJobOfferFormProps } from './IHrJobOfferFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FontSizes } from '@fluentui/theme/lib/fonts';

export default class HrJobOfferForm extends React.Component<IHrJobOfferFormProps, {}> {
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
      </div>
    );
  }
}
