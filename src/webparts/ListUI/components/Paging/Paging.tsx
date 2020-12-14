import * as React from "react";
import { IPagingProps } from "./IPagingProps";
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import Pagination from "react-js-pagination";
import styles from './Paging.module.scss';

export default class Paging extends React.Component<IPagingProps, null> {

  constructor(props: IPagingProps) {
    super(props);
    this._onBackToFirst = this._onBackToFirst.bind(this);
    this._onNextPage = this._onNextPage.bind(this);
  }

  public render(): React.ReactElement<IPagingProps> {
    return (
      <div>
        <DefaultButton onClick={this._onBackToFirst.bind(this)} text="back to first" />
        <PrimaryButton disabled={this.props.nextEnabled === undefined || this.props.nextEnabled === null ? true : (!this.props.nextEnabled)}  onClick={this._onNextPage.bind(this)} text="next" />
      </div>
    );
  }

  private _onBackToFirst(): void {
    this.props.onBackToFirst();
  }
  private _onNextPage(): void {
    this.props.onNextPage(this.props.currentPage + 1);
  }
}
