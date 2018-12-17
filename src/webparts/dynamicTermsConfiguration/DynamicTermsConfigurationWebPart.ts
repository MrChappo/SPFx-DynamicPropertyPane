import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneGroup,
  IPropertyPaneField
} from '@microsoft/sp-webpart-base';

import { Session, ITermSet, ITermSetData, ITerm, ITermData } from '@pnp/sp-taxonomy';
import { PropertyFieldColorPickerMini } from 'sp-client-custom-fields/lib/PropertyFieldColorPickerMini';
import { PropertyFieldTermPicker, IPickerTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';

import DynamicTermsConfiguration from './components/DynamicTermsConfiguration';
import { IDynamicTermsConfigurationProps } from './components/IDynamicTermsConfigurationProps';
import { IPropertyFieldColorPickerProps } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface ITermModel {
  id: string;
  name: string;
  color: string;
}

export interface IDynamicTermsConfigurationWebPartProps {
  terms: ITermModel[];
  termSet: IPickerTerms;
  termSetId: string;
  termSetSelected: boolean;
}

export default class DynamicTermsConfigurationWebPart extends BaseClientSideWebPart<IDynamicTermsConfigurationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDynamicTermsConfigurationProps> = React.createElement(
      DynamicTermsConfiguration,
      {
        terms: this.properties.terms,
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
    let labelsPropertiesGroup: IPropertyPaneGroup = {
      groupName: 'Content labels',
      groupFields: [
        PropertyFieldTermPicker('termSet', {
          allowMultipleSelections: false,
          context: this.context,
          excludeSystemGroup: true,
          initialValues: this.properties.termSet,
          isTermSetSelectable: true,
          key: 'termSetsPickerFieldId',
          label: 'Term set',
          onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => this.changeLabelTermSet(newValue),
          panelTitle: 'Select a term set for labels',
          properties: this.properties,
        })
      ]
    };

    if (this.properties.termSetSelected) {
      let colorPickers = this.constructColorPickers();
      if (colorPickers) {
        colorPickers.forEach(colorPicker => {
          labelsPropertiesGroup.groupFields.push(colorPicker);
        });
      }
    }

    return {
      pages: [
        {
          groups: [
            labelsPropertiesGroup
          ]
        }
      ]
    };
  }

  private changeLabelTermSet(input: IPickerTerms) {
    if (input.length === 0) {
      this.properties.termSetSelected = false;
      return null;
    } else {
      this.properties.termSetSelected = true;
      const termSetId: string = input[0]['key'];
      this.properties.termSetId = termSetId;
      this.properties.termSet = input;
      this.retrieveTermSet(termSetId);
    }
  }

  private retrieveTermSet(termSetId: string) {
    return new Promise<any>((resolve: () => void, reject: (error: any) => void): void => {
      const Taxonomy: Session = new Session(this.context.pageContext.site.absoluteUrl);
      const TermStore = Taxonomy.getDefaultKeywordTermStore();
      const TermSet: ITermSet = TermStore.getTermSetById(termSetId);
      TermSet.get().then((termSet: (ITermSetData & ITermSet)) => {
        termSet.terms.get().then((termsData: (ITermData & ITerm)[]) => {
          let terms: ITermModel[] = [];
          termsData.forEach((term: (ITermData & ITerm)) => {
            terms.push({
              id: term.Id.substring(6, 42),
              name: term.Name,
              color: term.LocalCustomProperties['labelColor']
            });
          });
          this.properties.terms = terms;
          if (this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
          }
          resolve();
        }).catch((error) => {
          console.error(error);
          reject(error);
        });
      });
    });
  }

  private constructColorPickers(): IPropertyPaneField<IPropertyFieldColorPickerProps>[] {
    if (this.properties.termSetSelected) {
      if (this.properties.terms !== undefined && this.properties.terms !== null) {
        let colorPickers: IPropertyPaneField<IPropertyFieldColorPickerProps>[] = [];
        this.properties.terms.forEach(term => {
          colorPickers.push(PropertyFieldColorPickerMini(term.id, {
            label: term.name,
            initialColor: term.color,
            onPropertyChange: (propertyPath: string, oldValue: string, newValue: string) => this.updateLabelColorField(term.id, newValue),
            render: this.render.bind(this),
            disableReactivePropertyChanges: this.disableReactivePropertyChanges,
            properties: this.properties,
            key: term.id
          }));
        });
        return colorPickers;
      } else {
        this.retrieveTermSet(this.properties.termSetId);
        return [];
      }
    } else {
      return [];
    }
  }

  private updateLabelColorField(labelId: string, labelColor: string): void {
    for (let i = 0; i < this.properties.terms.length; i++) {
      if (this.properties.terms[i].id === labelId) {
        this.properties.terms[i].color = labelColor;
      }
    }
  }
}
