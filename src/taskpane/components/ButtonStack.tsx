import * as React from 'react';
import { DefaultButton, PrimaryButton, Stack, IStackTokens } from 'office-ui-fabric-react';

export interface IButtonProps {
  // These are set based on the toggles shown above the examples (not needed in real code)
  disabled?: boolean;
  checked?: boolean;
  text: string;
}

// Example formatting
const stackTokens: IStackTokens = { childrenGap: 40 };

export const ButtonStack = (props:IButtonProps) => {
  const { disabled, checked,text } = props;

  return (
    <Stack tokens={stackTokens}>
      <DefaultButton text={text} onClick={_alertClicked} allowDisabledFocus disabled={disabled} checked={checked} />
      <PrimaryButton text={text} onClick={_alertClicked} allowDisabledFocus disabled={disabled} checked={checked} />
    </Stack>
  );
};

function _alertClicked(): void {
  alert('Clicked');
}
