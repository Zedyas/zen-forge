import { Popover } from '@base-ui/react/popover'
import { Plus, Signature, Trash2 } from 'lucide-react'
import { icon, smallIcon } from '../../ui/icons'
import { Tip } from '../../ui/Tip'
import { removeSignature, useSignatureStore } from './signatures'
import { setTool, toolKeys, useToolStore } from './tool-store'

/** The signature button: pick a saved signature to place, or draw a new one. */
export function SignatureTool() {
  const tool = useToolStore(state => state.tool)
  const signatures = useSignatureStore(state => state.signatures)

  const arm = (id: string): void => {
    setTool('signature')
    useToolStore.setState({ armedSignatureId: id })
  }

  // With no saved signature there is nothing to pick, so the button goes straight to drawing one.
  if (signatures.length === 0) {
    return (
      <Tip label="Signature" shortcut={toolKeys.signature}>
        <button type="button" className="tool" aria-label="Signature" onClick={() => useToolStore.setState({ signatureDialogOpen: true })}>
          <Signature {...icon} />
        </button>
      </Tip>
    )
  }

  return (
    <Popover.Root>
      <Tip label="Signature" shortcut={toolKeys.signature}>
        <Popover.Trigger className="tool" aria-label="Signature" aria-pressed={tool === 'signature'}>
          <Signature {...icon} />
        </Popover.Trigger>
      </Tip>
      <Popover.Portal>
        <Popover.Positioner sideOffset={6} align="start">
          <Popover.Popup className="popover pdf-signature-menu">
            <div className="popover-label">Place a signature</div>
            <ul>
              {signatures.map(signature => (
                <li key={signature.id}>
                  <Popover.Close className="pdf-signature-choice" onClick={() => arm(signature.id)} aria-label="Place this signature">
                    <img src={signature.dataUrl} alt="" />
                  </Popover.Close>
                  <Tip label="Delete signature">
                    <button type="button" className="tool" aria-label="Delete signature" onClick={() => removeSignature(signature.id)}>
                      <Trash2 {...smallIcon} />
                    </button>
                  </Tip>
                </li>
              ))}
            </ul>
            <Popover.Close className="button is-quiet pdf-signature-add" onClick={() => useToolStore.setState({ signatureDialogOpen: true })}>
              <Plus {...smallIcon} />Add signature…
            </Popover.Close>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
}
