import * as React from 'react';
import styles from './CustomColors.module.scss';
import { ICustomColorsProps } from './ICustomColorsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CustomColors extends React.Component<ICustomColorsProps, {}> {
  public render(): React.ReactElement<ICustomColorsProps> {
    let customColors = null;
    if(this.props.customColorsEnabled) {
      customColors = {
        'background-color': this.props.backgroundColor,
        'color': this.props.fontColor
      };
    }

    return (
      <div className={ styles.customColors } style={customColors}>
        <h3>Custom colors web part</h3>
        <p>Change the configuration to change the colors</p>
      </div>
    );
  }
}
