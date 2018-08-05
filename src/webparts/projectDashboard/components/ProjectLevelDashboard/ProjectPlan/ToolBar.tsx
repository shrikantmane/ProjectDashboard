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
 {/* {zoomRadios} */}
    return (
      <div className="zoom-bar">
        <ButtonToolbar>
          <Button className="zoom-bar-button" value="Days" onClick={this.handleZoomChange}>Days</Button>
          <Button className="zoom-bar-button" value="Months" onClick={this.handleZoomChange}>Months</Button>
        </ButtonToolbar>
      </div>
    );
  }
}
