angular.module('app', []);

angular.module('app').controller("MainController", function () {
    var vm = this;
    vm.title = 'AngularJS Practicle Example';
    vm.searchInput = '';
    vm.shows = [{
            title: '2 States',
            author: 'Chetan Bhagat',
            favorite: false
        },
        {
            title: 'Life of Pi',
            author: 'Davan',
            favorite: true
        },
        {
            title: 'Wings of Fire',
            author: 'APJ Abdul Kalam',
            favorite: true
        },
        {
            title: 'I too have a gf',
            author: 'Chetan Bhagat',
            favorite: false
        },
        {
            title: 'The Sealed Nector',
            author: 'Raheeq',
            favorite: true
        }
    ];
    vm.orders = [{
            id: 1,
            title: 'Author Ascending',
            key: 'author',
            reverse: false
        },
        {
            id: 2,
            title: 'Author Descending',
            key: 'author',
            reverse: true
        },
        {
            id: 3,
            title: 'Title Ascending',
            key: 'title',
            reverse: false
        },
        {
            id: 4,
            title: 'Title Ascending',
            key: 'title',
            reverse: true
        }
    ];
    vm.order = vm.orders[0];
    vm.new = {};
    vm.addShow = function () {
        vm.shows.push(vm.new);
        vm.new = {};
    };
});