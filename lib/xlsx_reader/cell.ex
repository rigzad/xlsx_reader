defmodule XlsxReader.Cell do
  @moduledoc """
  Cell structure.

  This structure contains the information of a cell in a sheet.

  - `value` - The value of the cell
  - `formula` - The formula used in the cell, if any
  - `ref` - The cell reference, like 'A1', 'B2', etc.
  - `bg_color` - The cell background color.

  This structure is used when the `cell_data_format` option is set to `:cell`.
  """

  defstruct [:value, :formula, :ref, :bg_color]

  @typedoc """
  XLSX cell data
  """
  @type t :: %__MODULE__{
          value: term(),
          formula: String.t() | nil,
          ref: String.t(),
          bg_color: String.t()
        }
end
