defmodule XlsxReader.Parsers.StylesParser do
  @moduledoc false

  # Parses SpreadsheetML style definitions.
  #
  # It extracts the relevant subset of style definitions in order to build
  # a `style_types` array which is used to look up the cell value format
  # for type conversions.

  @behaviour Saxy.Handler

  alias XlsxReader.{Array, Styles}
  alias XlsxReader.Parsers.Utils

  defmodule State do
    @moduledoc false
    defstruct pointer: nil,
              style_types: [],
              custom_formats: %{},
              fills: nil,
              supported_custom_formats: []
  end

  @spec parse(binary()) ::
          {:error, Saxy.ParseError.t()} | {:halt, any(), binary()} | {:ok, any(), any()}
  def parse(xml, supported_custom_formats \\ []) do
    with {:ok, state} <-
           Saxy.parse_string(xml, __MODULE__, %State{
             supported_custom_formats: supported_custom_formats
           }) do
      {:ok, state.style_types, state.custom_formats}
    end
  end

  @impl Saxy.Handler
  def handle_event(:start_document, _prolog, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_document, _data, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"numFmt", attributes}, state) do
    num_fmt_id = Utils.get_attribute(attributes |> IO.inspect, "numFmtId")
    format_code = Utils.get_attribute(attributes, "formatCode")
    {:ok, %{state | custom_formats: Map.put(state.custom_formats, num_fmt_id, format_code)}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"fills", _attributes}, state) do
    {:ok, %{state | pointer: {:collect_fills, {0, %{}}}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"fill", _attributes}, %{pointer: {:collect_fills, acc}} = state) do
    {:ok, %{state | pointer: {:collect_fill, nil, acc}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"cellXfs", _attributes}, state) do
    {:ok, %{state | pointer: :collect_xfs}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"xf", attributes}, %{pointer: :collect_xfs} = state) do
    num_fmt_id = Utils.get_attribute(attributes, "numFmtId")

    {:ok,
     %{
       state
       | style_types: [
           %{num_fmt: Styles.get_style_type(
             num_fmt_id,
             state.custom_formats,
             state.supported_custom_formats
           ), fill: state.fills[
            Utils.get_attribute(attributes, "fillId")
            |> String.to_integer()
            ]}
           | state.style_types
         ]
     }}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, {"bgColor", [{"rgb", c}]}, %{pointer: {:collect_fill, fill, fills}} = state) do
    {:ok, %{state | pointer: {:collect_fill, Map.put(fill || %{}, :bg_color, c), fills}}}
  end

  @impl Saxy.Handler
  def handle_event(:start_element, _element, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "cellXfs", state) do
    {:ok,
     %{state | pointer: nil, style_types: Array.from_list(Enum.reverse(state.style_types))}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "fill", %{pointer: {:collect_fill, fill, {i, fills}}} = state) do
    {:ok,
     %{state | pointer: {:collect_fills, {i + 1, Map.put(fills, i, fill)}}}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, "fills", %{pointer: {:collect_fills, {_, fills}}} = state) do
    {:ok,
     %{state | pointer: nil, fills: fills}}
  end

  @impl Saxy.Handler
  def handle_event(:end_element, _name, state) do
    {:ok, state}
  end

  @impl Saxy.Handler
  def handle_event(:characters, _chars, state) do
    {:ok, state}
  end
end
