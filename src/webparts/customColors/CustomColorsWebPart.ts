import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneField
} from '@microsoft/sp-webpart-base';

import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle, IPropertyFieldColorPickerProps } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import CustomColors from './components/CustomColors';
import { ICustomColorsProps } from './components/ICustomColorsProps';

export interface ICustomColorsWebPartProps {
  customColorsEnabled: boolean;
  backgroundColor: string;
  fontColor: string;
}

export default class CustomColorsWebPart extends BaseClientSideWebPart<ICustomColorsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICustomColorsProps > = React.createElement(
      CustomColors,
      {
        customColorsEnabled: this.properties.customColorsEnabled,
        backgroundColor: this.properties.backgroundColor,
        fontColor: this.properties.fontColor
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Styling configuration',
              groupFields: [
                PropertyFieldToggleWithCallout('customColorsEnabled', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'toggleCustomColors',
                  label: 'Use custom colors',
                  calloutContent: React.createElement('p', {}, 'Switch this to use custom styling'),
                  onText: 'ON',
                  offText: 'OFF',
                  checked: this.properties.customColorsEnabled
                }),
                ...this.customColorPickers(this.properties.customColorsEnabled)
              ]
            }
          ]
        }
      ]
    };
  }

  private customColorPickers(customColorsEnabled: boolean): IPropertyPaneField<IPropertyFieldColorPickerProps>[] {
    if (customColorsEnabled) {
      return [
        PropertyFieldColorPicker('backgroundColor', {
          label: 'Background color',
          selectedColor: this.properties.backgroundColor,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disabled: false,
          alphaSliderHidden: false,
          style: PropertyFieldColorPickerStyle.Full,
          iconName: 'Precipitation',
          key: 'backgroundColorFieldId'
        }),
        PropertyFieldColorPicker('fontColor', {
          label: 'Font color',
          selectedColor: this.properties.fontColor,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disabled: false,
          alphaSliderHidden: false,
          style: PropertyFieldColorPickerStyle.Full,
          iconName: 'Precipitation',
          key: 'fontColorFieldId'
        })
      ];
    } else {
      return [];
    }
  }
}
