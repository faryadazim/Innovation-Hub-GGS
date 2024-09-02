import * as React from "react";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import { IInnovationHubIntranetProps } from "./IInnovationHubIntranetProps";
import MainComponent from "./MainComponent";

export default class InnovationHubIntranet extends React.Component<
  IInnovationHubIntranetProps,
  {}
> {
  constructor(prop: IInnovationHubIntranetProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IInnovationHubIntranetProps> {
    return (
      <MainComponent
        context={sp}
        spcontext={this.props.context}
        graphContext={graph}
      />
    );
  }
}
