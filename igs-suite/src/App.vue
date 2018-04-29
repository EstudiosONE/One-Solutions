<template lang="pug">
#base
  <one-main>
  </one-main>
  one-icon-menu.icon-48.black-fill
</template>

<script>
import oneMain from "./modules/main/main.vue"
import oneIconMenu from "./components/icons/menu.vue"
import {
  DateRangePicker
} from '@syncfusion/ej2-calendars';

export default {
  name: 'app',
  components: {
    oneMain, oneIconMenu
  },
}
</script>

<style lang="scss">
body {
    margin: 0;
    position: absolute;
    width: 100%;
    min-height: 100%;
}
@import '/style/default'
</style>
