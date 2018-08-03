import * as React from 'react';

export default class Toolbar extends React.Component<any, any> {
  constructor(props) {
    super(props);
    this.handleZoomChange = this.handleZoomChange.bind(this);
  }

  handleZoomChange(e) {
    if(this.props.onZoomChange){
      this.props.onZoomChange(e.target.value)
    }
  }

  render() {
    let zoomRadios = ['Days', 'Months'].map((value) => {
      let isActive = this.props.zoom === value;
      return (
        <label key={value} className={`radioContainer radio-label ${isActive ? 'radio-label-active': ''}`}>
          <input type='radio'
             checked={isActive}
             onChange={this.handleZoomChange}
             value={value}/>
          {value}
          <span className="checkmark"></span>
        </label>
      );
    });

    return (
      <div className="zoom-bar">       
          {zoomRadios}
      </div>
    );
  } 
}
