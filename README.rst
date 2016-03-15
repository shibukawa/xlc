Excel Compiler
========================

Sample
---------------

* Formula Sample

  ``=(A2+B2)*C2*DATA-D2``

  * ``Data`` is a named range. It cell has integer ``10``.

* Output Sample

  .. code-block:: go

     package main

     func CalcDamage(power float64, weapon float64, hit float64, diffence float64) float64 {
         return (power+weapon)*hit*10 - diffence
     }

Current Feature
----------------------

* Create AST from Excel formula(https://github.com/shibukawa/xlsxformula)
* Parse range notation (like ``B2:D4``)(https://github.com/shibukawa/xlsxrange)
* Convert to Golang function(operator, paren, named range, range, function call)
* Getting value from named range and range

To Do
----------

* Convert function name to host language's API or add polyfill
* Add other language generation

Schema Definition
------------------------

This tool assumes the sheet named ``Schema`` has naiming rules to generate code.

.. list-table::
   :header-rows: 1

   - * Sheet Name
     * Define
     * Range
     * Name
   - * Attack
     * function
     * E2
     * calcDamage
   - *
     * param
     * A2
     * attack
   - *
     *
     * B2
     * weapon
   - *
     *
     * C2
     * hit
   - *
     *
     * D2
     * diffence


