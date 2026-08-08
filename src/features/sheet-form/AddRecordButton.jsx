import React from 'react';
import { Button } from '@univerjs/design';

/**
 * App-chrome trigger (not Univer's internal toolbar) that opens the Add Record panel.
 *
 * @param {{ onClick: () => void, disabled?: boolean, children?: React.ReactNode }} props
 */
export default function AddRecordButton({ onClick, disabled = false, children = 'Add Record' }) {
  return (
    <Button
      type="button"
      variant="default"
      size="small"
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </Button>
  );
}
