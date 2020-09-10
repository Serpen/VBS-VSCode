import KEYWORDS from "./keywords"
import MAIN_FUNCTIONS from "./mainFunctions"
import CONSTANTS from "./constants"
import TYPES from "./types"
import OPERATORS from "./operators"
import PROPS from './props'

const completions = [
    ...KEYWORDS,
    ...MAIN_FUNCTIONS,
    ...CONSTANTS,
    ...TYPES,
    ...OPERATORS,
    ...PROPS,
]

export default completions
