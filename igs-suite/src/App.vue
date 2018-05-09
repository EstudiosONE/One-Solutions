<template lang="pug">
  #App.o-app
    #Base.o-base
      one-header.o-header
      one-menu.o-menu
      router-view.o-navigator
      one-footer.o-footer
    #PopUp
</template>

<script>
export default {
  name: 'app',
  components: {
  },
}
</script>

<style lang="scss">
a:-webkit-any-link{
  text-decoration: none;
}
</style>
