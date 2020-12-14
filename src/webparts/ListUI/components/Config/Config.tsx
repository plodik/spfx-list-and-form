import * as React from 'react';
import { Fabric } from 'office-ui-fabric-react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IConfigProps } from './IConfigProps';
import styles from './Config.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export class Config extends React.Component<IConfigProps, {}> {

  constructor(props: IConfigProps) {
    super(props);
    this._handleBtnClick = this._handleBtnClick.bind(this);
  }
  public render(): JSX.Element {
    return (
      <Fabric>
        {this.props.displayMode === DisplayMode.Edit &&
          <div className={`${styles.placeholder}`}>
            <div className={styles.placeholderContainer}>
              <div className={styles.placeholderHead}>
                <div className={styles.placeholderHeadContainer}>
                  <i className={`${styles.placeholderIcon} ms-fontSize-su ms-Icon ms-ICon--CheckboxComposite`}></i>
                  <span className={`${styles.placeholderText} ms-fontWeight-light ms-fontSize-xxl`}>Begin with configuration...</span>
                </div>
              </div>
              <div className={styles.placeholderDescription}>
                <span className={styles.placeholderDescriptionText}>Please configure the web part</span>
              </div>
              <div className={styles.placeholderDescription}>
                <PrimaryButton
                  text="Configure"
                  ariaLabel="Configure"
                  ariaDescription="Please configure the web part"
                  onClick={this._handleBtnClick} />
              </div>
            </div>
          </div>
        }
        {this.props.displayMode === DisplayMode.Read &&
          <div className={`${styles.placeholder}`}>
            <div className={styles.placeholderContainer}>
              <div className={styles.placeholderHead}>
                <div className={styles.placeholderHeadContainer}>
                  <i className={`${styles.placeholderIcon} ms-fontSize-su ms-Icon ms-ICon--CheckboxComposite`}></i>
                  <span className={`${styles.placeholderText} ms-fontWeight-light ms-fontSize-xxl`}>Begin with configuration...</span>
                </div>
              </div>
              <div className={styles.placeholderDescription}>
                <span className={styles.placeholderDescriptionText}>Please configure the web part</span>
              </div>
            </div>
          </div>
        }
      </Fabric>
    );
  }
  private _handleBtnClick(event?: React.MouseEvent<HTMLButtonElement>) {
    this.props.configure();
  }
}
