import * as React from 'react';
import { ButtonGroup, Button, ButtonToolbar } from 'react-bootstrap';

export default class Toolbar extends React.Component<any, any> {
  constructor(props) {
    super(props);
    this.handleZoomChange = this.handleZoomChange.bind(this);
  }

  handleZoomChange(e) {
    if (this.props.onZoomChange) {
      this.props.onZoomChange(e.target.value)
    }
  }

  render() {
    let zoomRadios = ['Days', 'Months'].map((value) => {
      let isActive = this.props.zoom === value;
      return (
        <label key={value} className={`radioContainer radio-label ${isActive ? 'radio-label-active' : ''}`}>
          <input type='radio'
            checked={isActive}
            onChange={this.handleZoomChange}
            value={value} />
          {value}
          <span className="checkmark"></span>
        </label>
      );
    });
    {/* {zoomRadios} */ }
    let dayClass = "";
    let monthsClass = "";
    if (this.props.zoom == "Days") {
      dayClass = "zoom-bar-button zoom-button-active";
      monthsClass = "zoom-bar-button";
    } else {
      dayClass = "zoom-bar-button";
      monthsClass = "zoom-bar-button zoom-button-active";
    }
    return (
      <div className="zoom-bar">
        <ButtonToolbar>
          <Button className={dayClass} value="Days" onClick={this.handleZoomChange}>Days</Button>
          <Button className={monthsClass} value="Months" onClick={this.handleZoomChange}>Months</Button>
          <Button className="zoom-bar-button">Import</Button>
          <Button className="zoom-bar-button">View in Microsoft Project</Button>
        </ButtonToolbar>          
      </div>
    );
  }
}
