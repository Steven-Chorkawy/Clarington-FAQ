import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { Accordion } from '@pnp/spfx-controls-react';
import { getSiteSP } from '../../../pnpjs-config';

export default class FaqAccordion extends React.Component<IFaqAccordionProps, any> {

  constructor(props: any) {
    super(props);
    console.log('ctor');
    console.log(props);
    this.state = {
      items: undefined
    };
    getSiteSP().web.lists.getByTitle(this.props.listName).items().then(value => {
      this.setState({ items: value });
    });
  }

  public render(): React.ReactElement<IFaqAccordionProps> {
    // let sampleItems: any = [
    //   { Question: "Q1", Response: "R1" },
    //   { Question: "Q2", Response: "R2" },
    //   { Question: "Q3", Response: "R3" },
    //   { Question: "Q4", Response: "R4" },
    //   { Question: "Q5", Response: "R5" },
    //   { Question: "Q6", Response: "R6" },
    // ];
    if (this.state.items === undefined) {
      return <div>Loading...</div>;
    }
    else {
      return (
        <div>
          <p>{this.props.siteUrl} and {this.props.listName}.  {this.state.items.length} items found!</p>
          {this.state.items.map((item: any, index: number) => (
            <Accordion title={item[this.props.questionFieldName]} defaultCollapsed={true} className={"itemCell"} key={index}>
              <div className={"itemContent"}>
                <div className={"itemResponse"}>{item[this.props.answerFieldName]}</div>
              </div>
            </Accordion>
          ))}
        </div>
      );
    }
  }
}
