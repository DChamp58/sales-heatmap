// Minimal type declarations for the `leaflet.heat` plugin, which ships no types.
import 'leaflet';

declare module 'leaflet' {
  interface HeatMapOptions {
    minOpacity?: number;
    maxZoom?: number;
    max?: number;
    radius?: number;
    blur?: number;
    gradient?: Record<number, string>;
  }

  type HeatLatLngTuple = [number, number, number?];

  interface HeatLayer extends Layer {
    setLatLngs(latlngs: HeatLatLngTuple[]): this;
    addLatLng(latlng: HeatLatLngTuple): this;
    setOptions(options: HeatMapOptions): this;
    redraw(): this;
  }

  function heatLayer(
    latlngs: HeatLatLngTuple[],
    options?: HeatMapOptions
  ): HeatLayer;
}
