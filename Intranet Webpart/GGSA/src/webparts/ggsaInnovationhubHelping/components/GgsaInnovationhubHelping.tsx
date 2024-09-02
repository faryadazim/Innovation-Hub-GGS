import * as React from 'react'; 
import { IGgsaInnovationhubHelpingProps } from './IGgsaInnovationhubHelpingProps'; 
import MainComponents from './MainComponents';

export default class GgsaInnovationhubHelping extends React.Component<IGgsaInnovationhubHelpingProps, {}> {
  public render(): React.ReactElement<IGgsaInnovationhubHelpingProps> {
   console.log(this.props);

    return (
     <MainComponents/>
    );
  }
}
