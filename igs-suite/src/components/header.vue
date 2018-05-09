<template lang="pug">
#header
  .o-user-header
    .img
      one-icons-circle-user-male(v-if="profilePicture == '' & sex === 'M'")
      one-icons-circle-user-famale(v-if="profilePicture == '' & sex === 'F'")
      img(:src="profilePicture", v-if="profilePicture != ''")
    h1 {{ user }}
    h2 {{ position }}
    one-button-squad-48.icon-s
      one-icons-menu-vertical(slot="icon")
</template>

<script>
export default {
  name: "one-header",
  computed: {
    user: function() {
      const user = this.$store.getters.getUser
      return `${user.name} ${user.surname}`
    },
    position: function() {
      const user = this.$store.getters.getUser
      return user.position
    },
    profilePicture: function() {
      const user = this.$store.getters.getUser
      return user.profilePicture
    },
    sex: function(){
      const user = this.$store.getters.getUser
      return user.sex
    }
  }

}
</script>

<style lang="scss">
#header {
    display: grid;
    grid-template-rows: 48px;
    grid-template-columns: 1fr 248px;
    .o-user-header {
        grid-area: 1 / 2 / 2 / 3;
    }
}
</style>
