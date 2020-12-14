import * as React from 'react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import styles from './DatePickerRequired.module.scss';

const DayPickerStrings: IDatePickerStrings = {
  months: ['Leden', 'Únor', 'Březen', 'Duben', 'Květen', 'Červen', 'Červenec', 'Srpen', 'Září', 'Říjen', 'Listopad', 'Prosinec'],
  shortMonths: ['Led', 'Únr', 'Bře', 'Dub', 'Kvě', 'Čer', 'Čec', 'Srp', 'Zář', 'Říj', 'Lis', 'Pro'],
  days: ['neděle', 'pondělí', 'úterý', 'středa', 'čtvrtek', 'pátek', 'sobota'],
  shortDays: ['N', 'P', 'U', 'S', 'Č', 'P', 'S'],
  goToToday: 'Přejít na dnešek',
  prevMonthAriaLabel: 'Přejít na předcházející měsíc',
  nextMonthAriaLabel: 'Přejít na další měsíc',
  prevYearAriaLabel: 'Přejít na předcházející rok',
  nextYearAriaLabel: 'Přejít na další rok',
  isRequiredErrorMessage: 'Pole je povinné.',
  invalidInputErrorMessage: 'Neplatný formát datumu.'
};

export interface IDatePickerRequiredState {
  firstDayOfWeek?: DayOfWeek;
  selectedDate?: Date
}

export interface IDatePickerRequiredProps {
  onChange?: (selectedDate: Date) => void;
  minDate?: Date;
}

export class DatePickerRequired extends React.Component<IDatePickerRequiredProps, IDatePickerRequiredState> {
  constructor(props: {}) {
    super(props);
    this.state = { firstDayOfWeek: DayOfWeek.Monday };
  }

  public render(): JSX.Element {
    const { firstDayOfWeek } = this.state;

    return (
      <div className="docs-DatePicker">
        <DatePicker
          firstDayOfWeek={firstDayOfWeek}
          strings={DayPickerStrings}
          value={this.state.selectedDate}
          onSelectDate={this._DatePickerOnChange.bind(this)}
          formatDate={this._onFormatDate}
          placeholder="choose date..."
          ariaLabel="choose date"
        />
      </div>
    );
  }

  private _onFormatDate = (date: Date): string => { return date.getDate() + '.' + (date.getMonth() + 1) + '.' + (date.getFullYear()); };

  private _DatePickerOnChange(date: Date) {
    this.setState({ selectedDate: date });
    if (this.props.onChange) { this.props.onChange(date); }
  }
}
