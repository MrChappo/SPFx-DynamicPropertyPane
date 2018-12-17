import * as React from 'react';
import styles from './DynamicTermsConfiguration.module.scss';
import { IDynamicTermsConfigurationProps } from './IDynamicTermsConfigurationProps';

export default class DynamicTermsConfiguration extends React.Component<IDynamicTermsConfigurationProps, {}> {
  public render(): React.ReactElement<IDynamicTermsConfigurationProps> {
    return (
      <div className={styles.dynamicTermsConfiguration}>
        <h3 className={styles.header}>Dynamic terms web part</h3>
        <div>
          <h5>Term set: 'Labels'</h5>
          {this.showTerms()}
        </div>
      </div>
    );
  }

  private showTerms(): JSX.Element {
    let terms: JSX.Element[] = [];
    if (this.props && this.props.terms) {
      this.props.terms.forEach(term => {
        let colors = {
          "color": term.color
        };
        terms.push(<li style={colors}>{term.name}</li>);
      });
      return <ul>{terms}</ul>;
    } else {
      return <ul></ul>;
    }
  }
}
