import * as React from "react";

export interface ListItem {
  id: number;
  name: string;
  info: string;
  // icon: string;
}

export interface ListProps {
  message: string;
  items: ListItem[];
}

export default class List extends React.Component<ListProps> {
  render() {
    const { children, items, message } = this.props;

    const listItems = items.map((item) => (
      <li className="ms-ListItem" key={item.id}>
        {/* <i className={`ms-Icon ms-Icon--${item.icon}`}></i> */}
        {/* <span className="ms-font-m ms-fontColor-neutralPrimary">{item.primaryText}</span> */}
        <label htmlFor={item.name}>{item.name}:&nbsp;{item.info}</label><br />
        <input id={item.name} type="text" name={item.name} />
      </li>
    ))

    return (
      <main className="ms-welcome__main">
        {/* <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2> */}
        <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">{listItems}</ul>

        {children}
      </main>
    );
  }
}
