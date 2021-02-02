/* eslint-disable react/prop-types */
import * as React from "react";
import {
  Stack,
  MessageBar,
  MessageBarType,
  IStackProps
} from "office-ui-fabric-react";
import * as ReactDOM from "react-dom";

/* global document */

export function displayMessageBar(messageText: string) {
  interface IMessageBarProps {
    messageText?: string;
    visible?: boolean;
    onDismiss?: () => void;
  }

  const horizontalStackProps: IStackProps = {
    horizontal: true,
    tokens: { childrenGap: 16 }
  };
  const verticalStackProps: IStackProps = {
    styles: { root: { overflow: "hidden", width: "100%" } },
    tokens: { childrenGap: 20 }
  };

  const MBError = ({messageText, onDismiss}:IMessageBarProps) => (
    <MessageBar
      messageBarType={MessageBarType.error}
      isMultiline={false}
      onDismiss={onDismiss}
      dismissButtonAriaLabel="Close"
    >
      {messageText}
    </MessageBar>
  );

  const BasicMessageBar: React.FunctionComponent<IMessageBarProps> = ({ messageText = "PlaceholderText"}) => {
    
    const [visible, setVisible] = React.useState<boolean | undefined>(true);

    const setVisibility = React.useCallback((visible: boolean) => setVisible(visible), []);

    const props: IMessageBarProps = { messageText, visible: true, onDismiss: () => setVisibility(false) };

    return visible !== false ? (
      <Stack {...horizontalStackProps}>
        <Stack {...verticalStackProps}>
          <MBError {...props} />
        </Stack>
      </Stack>
    ) : null;
  };

  ReactDOM.render(<BasicMessageBar messageText={messageText} />, document.getElementById("messagebar"));
}
