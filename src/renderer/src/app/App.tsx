import { CommandController } from './CommandController'
import { WindowScreen } from './WindowScreen'
import { TipProvider } from '../ui/Tip'

export function App() {
  return (
    <TipProvider>
      <WindowScreen />
      <CommandController />
    </TipProvider>
  )
}
