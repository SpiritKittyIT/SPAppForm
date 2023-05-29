export interface ILang{
  Webpart: {
    Properties: {
      Description: string,
      BasicGroupName: string,
      InputFieldLabel: string
    }
  },
  Extension: {
    ButtonTitles: {
      New: string,
      Edit: string,
      Display: string
    }
  }
}

export const getLangStrings = async (locale: string): Promise<ILang> => {
  switch (locale) {
    case "en":
      return await import(/* webpackChunkName: 'lang' */'../../common/lang/en.json')
    default:
      return await import(/* webpackChunkName: 'lang' */'../../common/lang/en.json')
  }
}
