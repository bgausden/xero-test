import * as React from "react";
/* import { Button, ButtonType } from "office-ui-fabric-react"; */
import Header from "./Header";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { ButtonStack } from "./ButtonStack";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import { PrimaryButton, Stack } from "office-ui-fabric-react";
import { csvOnload } from "../../csv-functions";
import { CSVPicker } from "./FilePicker";

/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        const file = document.getElementById("file") as HTMLInputElement;
        const reader = new FileReader();

        if (file?.files && file.files[0]) {
          reader.onload = csvOnload(reader, context);
          reader.readAsText(file.files[0]);
        }
      });
    } catch (error) {
      console.error(error);
    }
  };

  buttonPrimary = ButtonStack;

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets\Glow Corporate Vert Prot 300w.png"
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets\Glow Corp Vert.svg" title={this.props.title} message="HSBC Xero" />
        {/*         <HeroList message="Convert HSBC Txn CSV to Xero's format." items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          

        </HeroList> */}
        <CSVPicker/>
        <Stack
          tokens={{
            childrenGap: "m",
            padding: "m"
          }}
        >
          <span className="ms-font-l ms-fontColor-neutralPrimary">Convert HSBC Txn CSV to Xeros format.</span>

          <p className="ms-font-m">
            Select a CSV file, then click <b>Process</b>.
          </p>

          <PrimaryButton text="Process HSBC CSV" onClick={() => console.log("clicked")} />
        </Stack>
      </div>
    );
  }
}
