import * as React from 'react'
//import styles from './appForm.module.scss'
import { escape } from '@microsoft/sp-lodash-subset'
import { GraphFI } from '@pnp/graph'
import { SPFI } from '@pnp/sp'
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown'

export interface IAppFormProps {
  description: string
  isDarkTheme: boolean
  environmentMessage: string
  hasTeamsContext: boolean
  userDisplayName: string
  sp: SPFI
  graph: GraphFI
}

const  AppForm: React.FC<IAppFormProps> = (props) => {
  const [groupOptions, groupOptionsSet] = React.useState<IDropdownOption[]>([])
  
  React.useEffect(() => {
    props.graph.groups().then((groups) => {
      console.log(groups)
    }).catch((err) => {console.error(err)})

    props.graph.groups().then((groups) => {
      groupOptionsSet(
        groups.map((group) => {
          return {
            key: group.id,
            text: group.displayName
          }
        })
      )
    }).catch((err) => {console.error(err)})
  }, [props])

  return (
    <>
      <div>Web part property value: <strong>{escape(props.description)}</strong></div>
      <Dropdown label={'label 1'} options={groupOptions}/>
    </>
  )
}

export default AppForm
